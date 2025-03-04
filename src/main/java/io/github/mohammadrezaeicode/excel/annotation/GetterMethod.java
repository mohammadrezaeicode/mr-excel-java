package io.github.mohammadrezaeicode.excel.annotation;

import java.lang.annotation.*;

@Documented
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
public @interface GetterMethod {
    String method() default "";
}
