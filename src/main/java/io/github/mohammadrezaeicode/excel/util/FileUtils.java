package io.github.mohammadrezaeicode.excel.util;

import io.github.mohammadrezaeicode.excel.model.ImageInput;
import io.github.mohammadrezaeicode.excel.model.ImageOutput;

import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.*;
import java.nio.ByteBuffer;
import java.util.*;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

public class FileUtils {
    public static byte[] convertImageOutputToByteArray(ImageOutput imageInput) throws IOException {
        switch (imageInput.getImage().getType()){
            case BASE64:
                return Base64.getDecoder().decode(((String) imageInput.getImage().getData()).split(",")[1]);
            case BUFFERED_IMAGE:
                ByteArrayOutputStream baos = new ByteArrayOutputStream();
                ImageIO.write((BufferedImage)imageInput.getImage().getData(), imageInput.getType(), baos);
                byte[] bytes = baos.toByteArray();
                return bytes;
            case FILE:
                ByteArrayOutputStream baos1 = new ByteArrayOutputStream();
                BufferedImage image= ImageIO.read((File) imageInput.getImage().getData());
                ImageIO.write((BufferedImage)imageInput.getImage().getData(), imageInput.getType(), baos1);
                byte[] bytes1 = baos1.toByteArray();
                return bytes1;
            default:
                throw new Error("TYPE is invalid");

        }
    }
    public static String writeFile(String fileName,Map<String,String> items, Map<String, ImageOutput> images) throws IOException {
        String generatedFileName=fileName+".xlsx";
        FileOutputStream fos = new FileOutputStream(generatedFileName);
        ZipOutputStream zipOut = new ZipOutputStream(fos);
        Set<String> keys=items.keySet();
        for (String key:keys){
            ZipEntry zipEntry = new ZipEntry(key);
            zipOut.putNextEntry(zipEntry);
            zipOut.write(items.get(key).getBytes());
        }
        keys=images.keySet();
        for (String key:keys ){
            byte[] data=convertImageOutputToByteArray(images.get(key));
            ZipEntry zipEntry = new ZipEntry(key);
            zipOut.putNextEntry(zipEntry);
            zipOut.write(data);
        }
        zipOut.close();
        fos.close();
        return generatedFileName;
    }
    public static ByteBuffer bufferResponse(Map<String,String> items, Map<String, ImageOutput> images)throws IOException{
        return ByteBuffer.wrap(byteArrayResponse(items,images));}
    public static String base64Response(Map<String,String> items, Map<String, ImageOutput> images)throws IOException{
        return Base64.getEncoder().encodeToString(byteArrayResponse(items,images));
    }
    public static String binaryStringResponse(Map<String,String> items, Map<String, ImageOutput> images) throws IOException {
        StringBuilder binaryString = new StringBuilder();
        var zipData=byteArrayResponse(items,images);
        for (byte b : zipData) {
            int value = b & 0xFF;
            String binaryVal = Integer.toBinaryString(value);
            while (binaryVal.length() < 8) {
                binaryVal = "0" + binaryVal;
            }
            binaryString.append(binaryVal);
        }
        return binaryString.toString();
    }
    public static InputStream inputStreamResponse(Map<String,String> items, Map<String, ImageOutput> images)throws IOException{
        return new ByteArrayInputStream(byteArrayResponse(items,images));
    }
    public static byte[] byteArrayResponse(Map<String,String> items, Map<String, ImageOutput> images) throws IOException {
        ByteArrayOutputStream byteArrayOutputStream = new ByteArrayOutputStream();
        ZipOutputStream zipOut = new ZipOutputStream(byteArrayOutputStream);
        Set<String> keys=items.keySet();
        for (String key:keys){
            ZipEntry zipEntry = new ZipEntry(key);
            zipOut.putNextEntry(zipEntry);
            zipOut.write(items.get(key).getBytes());
        }
        keys=images.keySet();
        for (String key:keys ){
            byte[] data=convertImageOutputToByteArray(images.get(key));
            ZipEntry zipEntry = new ZipEntry(key);
            zipOut.putNextEntry(zipEntry);
            zipOut.write(data);
        }
        zipOut.close();
        return byteArrayOutputStream.toByteArray();
    }
}
