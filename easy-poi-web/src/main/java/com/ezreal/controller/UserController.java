package com.ezreal.controller;

import com.ezreal.model.UserDTO;
import com.ezreal.model.UserVO;
import com.ezreal.service.UserServiceImpl;
import com.github.dozermapper.core.DozerBeanMapperBuilder;
import com.github.dozermapper.core.Mapper;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;

import javax.annotation.Resource;

@RestController
public class UserController {

    @Resource
    private UserServiceImpl userService;


    @GetMapping("/api/web/user/generateExcel")
    public String getUserExcel() {



        //when
        //then
        return "";

    }

    public static void main(String[] args) {
        //given
        UserDTO userDTO = new UserDTO();
        userDTO.setUserName("憨憨");
        userDTO.setPassword("hh123");
        userDTO.setPhone("123");
        userDTO.setSex(0);

        //when
        Mapper mapper = DozerBeanMapperBuilder.buildDefault();
        UserVO userVO = mapper.map(userDTO, UserVO.class);
        //then
        System.out.println(userVO);

    }
}
