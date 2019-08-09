package com.sh.docresolving.service;

import com.baomidou.mybatisplus.service.impl.ServiceImpl;
import com.sh.docresolving.dao.ConvertRecordDao;
import com.sh.docresolving.dto.ExcelTransformDto;
import com.sh.docresolving.entity.ConvertRecordEntity;
import com.sh.docresolving.utils.Excel2Pdf;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.util.Assert;
import org.springframework.util.StringUtils;

import java.io.File;

@Service
public class ConvertRecordService extends ServiceImpl<ConvertRecordDao, ConvertRecordEntity>{

    @Autowired
    private ConvertRecordDao convertRecordDao;

    public boolean saveConvertRecordByExcel2Pdf(ExcelTransformDto excelTransformDto,String fastOutUrl){
        String fileIn = excelTransformDto.getFileIn();
        ConvertRecordEntity convertRecordEntity = new ConvertRecordEntity();
        convertRecordEntity.setOriginalUrl(excelTransformDto.getFileIn());
        convertRecordEntity.setOriginalSuffix(fileIn.substring(fileIn.lastIndexOf(".")+1));
        convertRecordEntity.setConvertedUrl(fastOutUrl);
        convertRecordEntity.setConvertedSuffix("pdf");
        convertRecordEntity.setFileName(System.currentTimeMillis()+"");
        return this.insert(convertRecordEntity);
    }
}
