package com.ezreal.util;

import java.lang.reflect.InvocationTargetException;
import java.util.concurrent.ExecutionException;

/**
 * 关于异常的工具类
 *
 * @author ezreal
 */
public class ExceptionUtil {

    /**
     * 将CheckedException转换为RuntimeException重新抛出
     */
    public static RuntimeException unchecked(Throwable t) {
        if (t instanceof RuntimeException) {
            throw (RuntimeException) t;
        }
        if (t instanceof Error) {
            throw (Error) t;
        }
        throw new UncheckedException(t);
    }

    /**
     * 如果是著名的包裹类,从cause中获得真正异常,其他异常则不变
     */
    public static Throwable unwrap(Throwable t) {
        if (t instanceof ExecutionException || t instanceof InvocationTargetException || t instanceof UncheckedException) {
            return t.getCause();
        }
        return t;
    }

    /**
     * 组合unchecked与unwrap的效果
     */
    public static RuntimeException uncheckedAndWrap(Throwable t) {
        Throwable unwrapped = unwrap(t);
        if (unwrapped instanceof RuntimeException) {
            throw (RuntimeException) unwrapped;
        }
        if (unwrapped instanceof Error) {
            throw (Error) unwrapped;
        }
        throw new UncheckedException(unwrapped);
    }

    /**
     * 自定义一个CheckedException的wrapper
     */
    public static class UncheckedException extends RuntimeException {

        private static final long serialVersionUID = 1097556215090641638L;

        public UncheckedException(Throwable cause) {
            super(cause);
        }

        @Override
        public String getMessage() {
            return super.getCause().getMessage();
        }
    }

}

