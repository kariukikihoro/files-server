package com.muriithi.filesserver.serve;

import java.time.Instant;

public class TokenInfo {
    final String filename;
    final String type;
    final Instant expiry;
    final String tokenHash;

    TokenInfo(String filename, String type, Instant expiry, String tokenHash) {
        this.filename = filename;
        this.type = type;
        this.expiry = expiry;
        this.tokenHash = tokenHash;
    }
}
