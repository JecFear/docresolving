package com.sh.docresolving.controller;

import com.sh.docresolving.service.FastDFSService;
import com.sh.docresolving.utils.FastDFSClient;
import io.swagger.annotations.Api;
import io.swagger.annotations.ApiOperation;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;
import sun.misc.BASE64Encoder;

import java.io.IOException;

@RestController
@Api(tags = "附件上传")
public class UploadController {

    @Autowired
    private FastDFSClient dfsClient;

    @Autowired
    private FastDFSService fastDFSService;

    @PostMapping("/upload/file")
    @ApiOperation(value = "上传附件FADS", notes = "上传附件FADS", response = ResponseEntity.class)
    public ResponseEntity<String> uploadS(@RequestParam("file") MultipartFile multipartFile) throws IOException {
        String fileUrl = fastDFSService.upload(multipartFile);
        return new ResponseEntity<>(fileUrl, HttpStatus.OK);
    }

    @PostMapping("/upload/file-by-file-path")
    @ApiOperation(value = "上传附件FADS", notes = "上传附件FADS", response = ResponseEntity.class)
    public ResponseEntity<String> uploadS(@RequestParam("filePath") String filePath) throws IOException {
        String fileUrl = fastDFSService.upload(filePath);
        return new ResponseEntity<>(fileUrl, HttpStatus.OK);
    }

    @PostMapping("/download/file")
    @ApiOperation(value = "下载文件", notes = "下载文件", response = ResponseEntity.class)
    public ResponseEntity<String> downloadFile(@RequestParam("url") String url,@RequestParam("filePath") String filePath) throws IOException {
        String fileUrl = fastDFSService.downloadFile(url,filePath);
        return new ResponseEntity<>(fileUrl, HttpStatus.OK);
    }


    /**
     * 上传返回图片地址
     *
     * @param file
     * @return
     * @throws Exception
     */
    @ResponseBody
    @RequestMapping(value = "/upload/image", method = RequestMethod.POST)
    public ResponseEntity upload(@RequestParam MultipartFile file) throws Exception {
        String imgUrl = dfsClient.uploadFile(file);
        return ResponseEntity.ok(imgUrl);
    }

    /**
     *
     * @param multipartFile
     * @return
     * @throws Exception
     */
    @ResponseBody
    @ApiOperation(value = "上传附件返回base64字符串", notes = "上传附件返回base64字符串", response = ResponseEntity.class)
    @RequestMapping(value = "/upload/base64", method = RequestMethod.POST)
    public ResponseEntity<String> uploadToBase64Img(@RequestParam("file") MultipartFile multipartFile) throws Exception {
        String base64Img = new BASE64Encoder().encode(multipartFile.getBytes());
        return new ResponseEntity<>(base64Img, HttpStatus.OK);
    }
}
