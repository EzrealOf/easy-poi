package com.ezreal.controller;

import com.ezreal.model.UserDTO;
import com.ezreal.model.UserVO;
import com.github.dozermapper.core.DozerBeanMapperBuilder;
import com.github.dozermapper.core.Mapper;

public class UserController {

    public String getUser() {

        //given
        UserDTO userDTO = new UserDTO();
        userDTO.setUserName("憨憨");
        userDTO.setPassword("hh123");
        userDTO.setPhone("123");
        userDTO.setSex(0);

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
