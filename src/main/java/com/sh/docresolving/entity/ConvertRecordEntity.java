package com.sh.docresolving.entity;

import com.baomidou.mybatisplus.annotations.TableId;
import com.baomidou.mybatisplus.annotations.TableName;
import com.baomidou.mybatisplus.enums.IdType;
import lombok.Data;

import java.io.Serializable;
import java.util.Date;

@Data
@TableName("dr_convert_record")
public class ConvertRecordEntity implements Serializable {
    private static final long serialVersionUID = 1L;

    @TableId(type= IdType.AUTO)
    private Integer id;

    private String fileName;

    private String originalUrl;

    private String originalSuffix;

    private String convertedUrl;

    private String convertedSuffix;

    private Integer convertType;

    private String operator;

    private Date operateTime;
}
