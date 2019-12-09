package com.sh.docresolving.controller;

import com.sh.docresolving.service.PdfResolvingService;
import io.swagger.annotations.Api;
import io.swagger.annotations.ApiOperation;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;

import java.util.List;

/**
 * @Author:Dawn
 * @Date: 2019/10/22 9:41
 **/
@Api("pdf处理API")
@RestController
@RequestMapping("/pdf-convert")
public class PdfResolvingController {

    @Autowired
    private PdfResolvingService pdfResolvingService;

    @ApiOperation(value = "pdf合并", notes = "itext", response = ResponseEntity.class)
    @RequestMapping(value = "/pdf-merge", method = RequestMethod.POST)
    public String excel2PdfFormat(@RequestParam("localPaths") List<String> localPaths) throws Exception{
        return pdfResolvingService.pdfMerge(localPaths);
    }
}
