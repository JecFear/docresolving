package com.sh.docresolving;

import org.mybatis.spring.annotation.MapperScan;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.cloud.openfeign.EnableFeignClients;

@SpringBootApplication
@MapperScan(basePackages = {"com.sh.docresolving.dao"})
@EnableFeignClients
public class DocresolvingApplication {

    public static void main(String[] args) {
        SpringApplication.run(DocresolvingApplication.class, args);
    }

}
