package com.muriithi.filesserver.serve;

import jakarta.annotation.PreDestroy;
import jakarta.servlet.http.HttpServletRequest;
import jakarta.servlet.http.HttpServletResponse;
import lombok.RequiredArgsConstructor;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Component;

import javax.crypto.Mac;
import javax.crypto.spec.SecretKeySpec;
import java.net.URLEncoder;
import java.nio.charset.StandardCharsets;
import java.security.NoSuchAlgorithmException;
import java.security.SecureRandom;
import java.time.Instant;
import java.time.temporal.ChronoUnit;
import java.util.Base64;
import java.util.Map;
import java.util.Set;
import java.util.concurrent.ConcurrentHashMap;
import java.util.concurrent.Executors;
import java.util.concurrent.ScheduledExecutorService;
import java.util.concurrent.TimeUnit;
import java.util.regex.Pattern;

@Component
@RequiredArgsConstructor
public class UtilityMethodsService {

    private final Map<String, TokenInfo> activeTokens = new ConcurrentHashMap<>();

    private final Map<String, RateLimitInfo> rateLimitMap = new ConcurrentHashMap<>();

    private final ScheduledExecutorService scheduler = Executors.newScheduledThreadPool(1);

    private final SecureRandom secureRandom = new SecureRandom();

    private static final int MAX_TOKENS_PER_MINUTE = 10;

    private static final Set<String> OFFICE_EXTENSIONS = Set.of(
            "doc", "docx", "xls", "xlsx", "ppt", "pptx", "rtf"
    );
    private static final Pattern PRIVATE_IP_PATTERN = Pattern.compile(
            "(^127\\.)|(^10\\.)|(^172\\.1[6-9]\\.)|(^172\\.2[0-9]\\.)|(^172\\.3[0-1]\\.)|(^192\\.168\\.)"
    );

    @Value("${server.public-url:#{null}}")
    private String publicUrl;

    @Value("${file.token.expiry.minutes:5}")
    private int tokenExpiryMinutes;

    @Value("${cors.allowed-origins:*}")
    private String allowedOrigins;

    @Value("${token.secret:default-secret}")
    private String tokenSecret;



    public String getClientIp(HttpServletRequest request) {
        String ip = request.getHeader("X-Forwarded-For");
        if (ip != null && !ip.isEmpty() && !"unknown".equalsIgnoreCase(ip)) {
            String[] ipArray = ip.split(",");
            if (ipArray.length > 0) {
                return ipArray[0].trim();
            }
        }
        ip = request.getHeader("Proxy-Client-IP");
        if (ip != null && !ip.isEmpty() && !"unknown".equalsIgnoreCase(ip)) {
            return ip;
        }
        ip = request.getHeader("WL-Proxy-Client-IP");
        if (ip != null && !ip.isEmpty() && !"unknown".equalsIgnoreCase(ip)) {
            return ip;
        }
        return request.getRemoteAddr();
    }

    private String getServerBaseUrl(HttpServletRequest request) {
        if (publicUrl != null && !publicUrl.isEmpty()) {
            return publicUrl;
        }

        String scheme = request.getScheme();
        String serverName = request.getServerName();
        int serverPort = request.getServerPort();
        String contextPath = request.getContextPath();

        if (("http".equals(scheme) && serverPort == 80) ||
                ("https".equals(scheme) && serverPort == 443)) {
            return scheme + "://" + serverName + contextPath;
        } else {
            return scheme + "://" + serverName + ":" + serverPort + contextPath;
        }
    }

    public String getPublicFileUrl(String type, String filename, HttpServletRequest request) {
        return getServerBaseUrl(request) + "/api/files/download?type=" +
                URLEncoder.encode(type, StandardCharsets.UTF_8) +
                "&filename=" + URLEncoder.encode(filename, StandardCharsets.UTF_8);
    }

    public String getPublicFileUrlWithToken(String filename, String token, HttpServletRequest request) {
        return getServerBaseUrl(request) + "/api/files/public-view?token=" +
                URLEncoder.encode(token, StandardCharsets.UTF_8) +
                "&filename=" + URLEncoder.encode(filename, StandardCharsets.UTF_8);
    }

    public String generateToken(String filename, String type) {
        byte[] randomBytes = new byte[32];
        secureRandom.nextBytes(randomBytes);
        String plainToken = Base64.getUrlEncoder().withoutPadding().encodeToString(randomBytes);
        String tokenHash = hashToken(plainToken);

        activeTokens.put(tokenHash, new TokenInfo(
                filename,
                type,
                Instant.now().plus(tokenExpiryMinutes, ChronoUnit.MINUTES),
                tokenHash
        ));
        return plainToken;
    }

    private String hashToken(String token) {
        try {
            Mac mac = Mac.getInstance("HmacSHA256");
            SecretKeySpec secretKeySpec = new SecretKeySpec(tokenSecret.getBytes(StandardCharsets.UTF_8), "HmacSHA256");
            mac.init(secretKeySpec);
            byte[] hash = mac.doFinal(token.getBytes(StandardCharsets.UTF_8));
            return Base64.getEncoder().encodeToString(hash);
        } catch (NoSuchAlgorithmException | java.security.InvalidKeyException e) {
            throw new RuntimeException("Failed to hash token", e);
        }
    }

    public TokenInfo validateToken(String token, String filename) {
        String tokenHash = hashToken(token);
        TokenInfo info = activeTokens.get(tokenHash);
        if (info == null || Instant.now().isAfter(info.expiry) || !filename.equals(info.filename)) {
            return null;
        }
        return info;
    }

    public boolean isRateLimited(String clientIp) {
        RateLimitInfo info = rateLimitMap.compute(clientIp, (key, existing) -> {
            if (existing == null) {
                return new RateLimitInfo();
            }

            if (existing.timestamp.plus(1, ChronoUnit.MINUTES).isBefore(Instant.now())) {
                return new RateLimitInfo();
            }

            existing.count.incrementAndGet();
            return existing;
        });

        return info.count.get() > MAX_TOKENS_PER_MINUTE;
    }

    public void cleanupExpiredTokens() {
        Instant now = Instant.now();
        activeTokens.entrySet().removeIf(entry -> now.isAfter(entry.getValue().expiry));
        rateLimitMap.entrySet().removeIf(entry ->
                entry.getValue().timestamp.plus(1, ChronoUnit.HOURS).isBefore(now)
        );
    }

    public boolean isValidInput(String type, String filename) {
        return isValidType(type) && isValidFilename(filename);
    }

    public boolean isValidType(String type) {
        return type != null && type.matches("^[a-zA-Z0-9_-]+$") && type.length() <= 50;
    }

    public boolean isValidFilename(String filename) {
        if (filename == null || filename.trim().isEmpty()) {
            return false;
        }

        if (filename.contains("..") || filename.contains("/") || filename.contains("\\")) {
            return false;
        }

        if (filename.chars().anyMatch(c -> c < 32 || c == 127)) {
            return false;
        }

        if (!filename.contains(".") || filename.endsWith(".")) {
            return false;
        }

        return filename.length() <= 255 &&
                filename.matches("^[\\w\\-.\\s()\\[\\]@$!]+\\.[A-Za-z0-9]+$");
    }

    public void setCorsHeaders(HttpServletResponse response) {
        response.setHeader("Access-Control-Allow-Origin", allowedOrigins);
        response.setHeader("Access-Control-Allow-Methods", "GET, HEAD, OPTIONS");
        response.setHeader("Access-Control-Allow-Headers", "Content-Type, Authorization");
        response.setHeader("Access-Control-Max-Age", "3600");
    }

    @PreDestroy
    public void cleanup() {
        activeTokens.clear();
        rateLimitMap.clear();
        if (scheduler != null && !scheduler.isShutdown()) {
            scheduler.shutdown();
            try {
                if (!scheduler.awaitTermination(5, TimeUnit.SECONDS)) {
                    scheduler.shutdownNow();
                }
            } catch (InterruptedException e) {
                scheduler.shutdownNow();
                Thread.currentThread().interrupt();
            }
        }
    }
}
