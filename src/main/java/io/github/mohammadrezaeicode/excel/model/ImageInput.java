package io.github.mohammadrezaeicode.excel.model;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.io.File;

@NoArgsConstructor
@AllArgsConstructor
@Data
@Builder
public class ImageInput {
    private AcceptType type;
    private ImageExtension extension;
    private Object data;

    public ImageInput(Object image, ImageExtension extension, AcceptType type) {
        if (AcceptType.FILE == type && !(image instanceof File)) {
            throw new Error("image should be File");

        }
        if (AcceptType.BASE64 == type && !(image instanceof String)) {
            throw new Error("image should be String(base64)");
        }
        this.data = image;
        this.type = type;
        this.extension = extension;
    }

    public enum ImageExtension {
        PNG("png"), JPEG("jpeg"), JPG("jpg"), GIF("gif");
        private final String name;

        ImageExtension(String name) {
            this.name = name;
        }

        public static ImageExtension getExtensionByName(String name) {
            switch (name) {
                case "jpeg":
                    return JPEG;
                case "jpg":
                    return JPG;
                case "gif":
                    return GIF;
                default:
                    return PNG;
            }
        }

        public String getName() {
            return name;
        }
    }

    public enum AcceptType {
        BASE64, FILE, URL, BUFFERED_IMAGE
    }
}
