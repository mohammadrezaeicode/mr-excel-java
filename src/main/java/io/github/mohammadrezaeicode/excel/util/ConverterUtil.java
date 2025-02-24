package io.github.mohammadrezaeicode.excel.util;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.ByteBuffer;
import java.util.Base64;

public class ConverterUtil {

    public static byte[] inputStreamToByteArray(InputStream inputStream) throws IOException {
        try (ByteArrayOutputStream baos = new ByteArrayOutputStream()) {
            byte[] buffer = new byte[1024];
            int bytesRead;
            while ((bytesRead = inputStream.read(buffer)) != -1) {
                baos.write(buffer, 0, bytesRead);
            }
            return baos.toByteArray();
        }
    }

    public static byte[] byteBufferToByteArray(ByteBuffer buffer) {
        return buffer.array();
    }

    public static byte[] byteBufferToByteArrayV2(ByteBuffer buffer) {
        byte[] byteArray = new byte[buffer.remaining()];
        buffer.get(byteArray);
        return byteArray;
    }

    public static byte[] base64ToByteArray(String base64String) {
        return Base64.getDecoder().decode(base64String);
    }

    public static byte[] binaryStringToByteArray(String binaryString) {
        String[] binaryValues = binaryString.trim().split(" ");
        byte[] byteArray = new byte[binaryValues.length];
        for (int i = 0; i < binaryValues.length; i++) {
            byteArray[i] = (byte) Integer.parseInt(binaryValues[i], 2);
        }
        return byteArray;
    }

    public static byte[] binaryStringToByteArrayV1(String binaryString) {
        String[] binaryValues = binaryString.trim().split(" ");
        byte[] byteArray = new byte[binaryValues.length];

        for (int i = 0; i < binaryValues.length; i++) {
            byteArray[i] = (byte) Integer.parseInt(binaryValues[i], 2);
        }
        return byteArray;
    }

    public static byte[] binaryStringToByteArray2(String binaryString) {
        String cleanedBinaryString = binaryString.replaceAll("\\s+", "");
        int length = cleanedBinaryString.length();

        if (length % 8 != 0) {
            throw new IllegalArgumentException("Binary string length must be a multiple of 8.");
        }

        byte[] byteArray = new byte[length / 8];

        for (int i = 0; i < length; i += 8) {
            String byteString = cleanedBinaryString.substring(i, i + 8);
            byteArray[i / 8] = (byte) Integer.parseInt(byteString, 2);
        }
        return byteArray;
    }
}

