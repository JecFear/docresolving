package com.sh.docresolving.service;

import com.github.tobato.fastdfs.domain.conn.Connection;
import com.github.tobato.fastdfs.domain.conn.FdfsWebServer;
import com.github.tobato.fastdfs.domain.fdfs.*;
import com.github.tobato.fastdfs.domain.proto.storage.DownloadByteArray;
import com.github.tobato.fastdfs.domain.proto.storage.DownloadCallback;
import com.github.tobato.fastdfs.exception.FdfsServerException;
import com.github.tobato.fastdfs.service.DefaultGenerateStorageClient;
import com.github.tobato.fastdfs.service.FastFileStorageClient;
import com.github.tobato.fastdfs.service.TrackerClient;
import org.apache.commons.fileupload.FileItem;
import org.apache.commons.fileupload.disk.DiskFileItem;
import org.apache.commons.io.FilenameUtils;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.util.ObjectUtils;
import org.springframework.web.multipart.MultipartFile;
import org.springframework.web.multipart.commons.CommonsMultipartFile;

import java.io.*;
import java.nio.file.Files;
import java.util.List;

@Service
public class FastDFSService {

    @Autowired
    private FastFileStorageClient storageClient;

    @Autowired
    private FdfsWebServer fdfsWebServer;


    public String upload(File file) throws IOException {
        if(file.exists()){
            DiskFileItem fileItem = new DiskFileItem("file", Files.probeContentType(file.toPath()), true, file.getName(), 100000000, file.getParentFile());
            OutputStream ous = fileItem.getOutputStream();
            MultipartFile multipartFile = new CommonsMultipartFile(fileItem);
            StorePath storePath = storageClient.uploadFile(multipartFile.getInputStream(),multipartFile.getSize(), FilenameUtils.getExtension(multipartFile.getOriginalFilename()),null);
            ous.flush();
            ous.close();
            return getResAccessUrl(storePath);
        }else{
            throw new FileNotFoundException("no file content found");
        }
    }

    public String upload(String filePath) throws  IOException{
        File file = new File(filePath);
        String url = upload(file);
        return url;
    }

    public String upload(MultipartFile file) throws IOException {
        if(!ObjectUtils.isEmpty(file)){
            StorePath storePath = storageClient.uploadFile(file.getInputStream(),file.getSize(), FilenameUtils.getExtension(file.getOriginalFilename()),null);
            return getResAccessUrl(storePath);
        }else{
            throw new FileNotFoundException("no file content found");
        }
    }

    public String uploadFile(File file) throws IOException{
        if(!ObjectUtils.isEmpty(file)){
            InputStream fis=new FileInputStream(file);
            try{
                StorePath storePath = storageClient.uploadFile(fis,file.length(), FilenameUtils.getExtension(file.getName()),null);
                return getResAccessUrl(storePath);
            }catch (FdfsServerException ex){
                throw new FileNotFoundException(ex.getMessage());
            }finally {
                fis.close();
            }
        }else{
            throw new FileNotFoundException("no file content found");
        }
    }

    public String downloadFile(String url,String filePath){
        StorePath storePath = getGroupNameByUrl(url);
        String fileName = url.substring(url.lastIndexOf("/")+1);
        String fileOut = filePath + File.separator + fileName;
        DownloadByteArray callback = new DownloadByteArray();
        byte[] content = storageClient.downloadFile(storePath.getGroup(), storePath.getPath(), callback);
        BufferedOutputStream bos = null;
        FileOutputStream fos = null;
        File file = null;
        try {
            File dir = new File(filePath);
            if(!dir.exists()&&dir.isDirectory()){//判断文件目录是否存在
                dir.mkdirs();
            }
            file = new File(fileOut);
            fos = new FileOutputStream(file);
            bos = new BufferedOutputStream(fos);
            bos.write(content);
            bos.flush();
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            if (bos != null) {
                try {
                    bos.close();
                } catch (IOException e1) {
                    e1.printStackTrace();
                }
            }
            if (fos != null) {
                try {
                    fos.close();
                } catch (IOException e1) {
                    e1.printStackTrace();
                }
            }
        }
        return fileOut;
    }

    private String getResAccessUrl(StorePath storePath) {
        String fileUrl = fdfsWebServer.getWebServerUrl()+ storePath.getFullPath();
        return fileUrl;
    }

    private StorePath getGroupNameByUrl(String url){
        StorePath storePath = new StorePath();
        String infoPath = url.substring(url.indexOf("group"));
        String groupName = infoPath.substring(0,infoPath.indexOf("/"));
        String path = infoPath.substring(infoPath.indexOf("/")+1);
        storePath.setGroup(groupName);
        storePath.setPath(path);
        return storePath;
    }
}
