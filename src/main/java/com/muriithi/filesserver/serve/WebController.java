package com.muriithi.filesserver.serve;


import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;

@Controller
public class WebController {

    @GetMapping("/")
    public String index() {
        return "forward:/index.html";
    }

    @GetMapping("/admin")
    public String admin() {
        return "forward:/index.html";
    }
}
