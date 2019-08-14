package com.sh.docresolving.service;

import com.baomidou.mybatisplus.service.impl.ServiceImpl;
import com.sh.docresolving.dao.ConvertRecordDao;
import com.sh.docresolving.dto.ExcelTransformDto;
import com.sh.docresolving.entity.ConvertRecordEntity;
import com.sh.docresolving.utils.Excel2Pdf;
import net.shouhouzn.lims.ms.umps.model.UserModel;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.util.Assert;
import org.springframework.util.StringUtils;

import java.io.File;
import java.util.Date;

@Service
public class ConvertRecordService extends ServiceImpl<ConvertRecordDao, ConvertRecordEntity>{

    @Autowired
    private ConvertRecordDao convertRecordDao;

    public boolean saveConvertRecordByExcel2Pdf(ExcelTransformDto excelTransformDto, String fastOutUrl, UserModel userModel){
        String fileIn = excelTransformDto.getFileIn();
        ConvertRecordEntity convertRecordEntity = new ConvertRecordEntity();
        convertRecordEntity.setOriginalUrl(excelTransformDto.getFileIn());
        convertRecordEntity.setOriginalSuffix(fileIn.substring(fileIn.lastIndexOf(".")+1));
        convertRecordEntity.setConvertedUrl(fastOutUrl);
        convertRecordEntity.setConvertedSuffix("pdf");
        convertRecordEntity.setFileName(System.currentTimeMillis()+"");
        Assert.isTrue(userModel.getUser()!=null&&userModel.getUser().getId()!=null,"未获取到当前登录人，请重新登录或联系管理员!");
        convertRecordEntity.setOperator(userModel.getUser().getId().toString());
        convertRecordEntity.setOperateTime(new Date());
        return this.insert(convertRecordEntity);
    }
}
