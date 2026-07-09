package com.example.demo;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.*;

import java.util.HashMap;
import java.util.Map;

@RestController
public class HelloController {

    @Autowired
    private UserRepository userRepository;

    @PostMapping("/login")
    public Map<String, Object> login(@RequestParam String username, @RequestParam String password) {
        Map<String, Object> result = new HashMap<>();

        if (username == null || username.trim().isEmpty() || password == null || password.trim().isEmpty()) {
            result.put("code", 1);
            result.put("message", "用户名和密码不能为空");
            return result;
        }

        User user = userRepository.findByUsername(username.trim());
        if (user == null) {
            result.put("code", 2);
            result.put("message", "账号不存在，请先注册");
            return result;
        }

        if (!user.getPassword().equals(password)) {
            result.put("code", 1);
            result.put("message", "密码错误，请重新输入");
            return result;
        }

        result.put("code", 0);
        result.put("message", "登录成功");
        result.put("username", user.getUsername());
        return result;
    }

    @PostMapping("/register")
    public Map<String, Object> register(@RequestBody Map<String, String> params) {
        Map<String, Object> result = new HashMap<>();
        String username = params.get("username");
        String password = params.get("password");

        if (username == null || username.trim().isEmpty() || password == null || password.trim().isEmpty()) {
            result.put("success", false);
            result.put("message", "用户名和密码不能为空");
            return result;
        }

        User existing = userRepository.findByUsername(username.trim());
        if (existing != null) {
            result.put("success", false);
            result.put("message", "用户名已存在，请更换");
            return result;
        }

        User user = new User(username.trim(), password);
        userRepository.save(user);
        result.put("success", true);
        result.put("message", "注册成功，请登录");
        return result;
    }
}
