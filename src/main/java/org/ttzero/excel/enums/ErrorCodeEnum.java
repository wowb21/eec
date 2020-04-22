/*
 * Copyright (c) 2019-2020, guanquan.wang@yandex.com All Rights Reserved.
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

package org.ttzero.excel.enums;

/**
 * 2.5.6 Error Codes
 * <p>
 * If the calculation of a formula results in an error
 * or any other action fails, Excel sets a specific error code.
 * These error codes are used for instance in cell records and formulas.
 *
 * @author guanquan.wang at 2019-01-25 14:18
 */
public enum ErrorCodeEnum {
    /**
     * Intersection of two cell ranges is empty
     */
    NULL(0, "#NULL!"),
    /**
     * Division by zero
     */
    DIV_ZERO(0x07, "#DIV/0!"),
    /**
     * Wrong type of operand
     */
    VALUE(0x0F, "#VALUE!"),
    /**
     * Illegal or deleted cell reference
     */
    REF(0x17, "#REF!"),
    /**
     * Wrong function or range name
     */
    NAME(0x1D, "#NAME?"),
    /**
     * Value range overflow
     */
    OVERFLOW(0x24, "#NUM!"),
    /**
     * Argument or function not available
     */
    NA(0x2A, "#N/A")
    ;
    String desc;
    byte code;

    ErrorCodeEnum(int code, String desc) {
        this.code = (byte) code;
        this.desc = desc;
    }

    public String getDesc() {
        return desc;
    }

    public int getCode() {
        return code;
    }

    public static ErrorCodeEnum of(byte code) {
        for (ErrorCodeEnum ec : values()) {
            if (ec.code == code) {
                return ec;
            }
        }
        return NULL;
    }
}
