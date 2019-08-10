package com.sh.docresolving.feign;

import com.sh.docresolving.dto.ExcelTransformDto;
import org.springframework.cloud.openfeign.FeignClient;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;

@FeignClient(name = "doc-resolving")
public interface FeignDocresolvingClient {

    @PostMapping("/excel-to-pdf")
    String Excel2Pdf(@RequestBody ExcelTransformDto excelTransformDto);
}
