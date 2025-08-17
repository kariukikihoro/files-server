package com.muriithi.filesserver.serve;

import java.time.Instant;
import java.util.concurrent.atomic.AtomicInteger;

public class RateLimitInfo {
    final AtomicInteger count = new AtomicInteger(1);
    final Instant timestamp = Instant.now();
}
