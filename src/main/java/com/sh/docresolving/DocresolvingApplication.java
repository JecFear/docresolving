package com.sh.docresolving;

import org.mybatis.spring.annotation.MapperScan;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

@SpringBootApplication
@MapperScan(basePackages = {"com.sh.docresolving.dao"})
public class DocresolvingApplication {

    public static void main(String[] args) {
        SpringApplication.run(DocresolvingApplication.class, args);
    }

}
