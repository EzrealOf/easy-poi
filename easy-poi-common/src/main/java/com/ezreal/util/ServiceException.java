package com.ezreal.util;

import lombok.Data;
import lombok.EqualsAndHashCode;
import lombok.NoArgsConstructor;

@EqualsAndHashCode(callSuper = true)
@Data
@NoArgsConstructor
public class ServiceException extends RuntimeException{

    private static final long serialVersionUID = 3615273440631030173L;

    private int status = 200;

    public ServiceException(String message) {
        super(message);
    }

    public ServiceException(String message, int status) {
        super(message);
        this.status = status;
    }

    public ServiceException(Throwable cause) {
        super(cause);
    }

    public ServiceException(String message, Throwable cause) {
        super(message, cause);
    }

    @Override
    public String toString() {
        return super.toString();
    }


}
