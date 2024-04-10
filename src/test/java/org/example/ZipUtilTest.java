package org.example;

import java.io.IOException;

import org.junit.jupiter.api.Test;

class ZipUtilTest {
    @Test
    void testZipFile() throws IOException {
        ZipUtil.zip("example.zip", "example.xlsx");
    
    }
}
