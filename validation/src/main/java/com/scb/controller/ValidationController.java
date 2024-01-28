package com.scb.controller;

import com.scb.service.ExcelValidatorService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;

@RestController
public class ValidationController {

    @Autowired
    ExcelValidatorService validatorService;

    @GetMapping("/validate")
    String validateReport() {
        return validatorService.validate();
    }

}
