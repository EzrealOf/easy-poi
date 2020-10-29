package com.ezreal.service;

import com.ezreal.model.UserDTO;
import com.ezreal.util.ExcelReadUtils;
import com.google.common.collect.Lists;
import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.stereotype.Service;

import java.util.List;

@Service
public class UserServiceImpl {

    private static final Integer SIZE = 100;

    public String generateUser(){
        //given
        List<UserDTO> userList = getUserList();
        Workbook workbook = ExcelReadUtils.getWorkbook("test.xml");

        return "";
    }

    private List<UserDTO> getUserList(){
        List<UserDTO> userDTOList = Lists.newArrayList();
        for (int i = 0; i < SIZE; i++) {
            UserDTO userDTO = new UserDTO();
            userDTO.setUserName("憨憨"+i+"号");
            userDTO.setPassword("hh123_"+i);
            userDTO.setPhone("123_"+i);
            userDTO.setSex(i%2);
            userDTOList.add(userDTO);
        }
        return userDTOList;

    }

}
