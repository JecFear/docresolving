package com.sh.docresolving.dto;

import lombok.Data;

/**
 * @Author:Dawn
 * @Date: 2019/10/19 17:23
 **/
@Data
public class PrintConvertDto {

    /**
     * 原始url
     */
    private String originalUrl;
    /**
     * 转换url
     */
    private String convertUrl;
    /**
     * 原始文件本地地址
     */
    private String originalPath;
    /**
     * 转换文件本地地址
     */
    private String convertPath;
}
