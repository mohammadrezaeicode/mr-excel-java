package io.github.mohammadrezaeicode.excel.model.types;

import java.io.InputStream;
import java.nio.ByteBuffer;

public enum GenerateType {
    NODE_BUFFER(byte[].class),BYTE_BUFFER(ByteBuffer.class),INPUT_STREAM(InputStream.class),BYTE_ARRAY(byte[].class) , ARRAY(byte[].class) , BINARY_STRING(String.class) , BASE64(String.class);
    private Class type;
    GenerateType(Class type){
        this.type=type;
    }

    public Class getType() {
        return type;
    }
}
