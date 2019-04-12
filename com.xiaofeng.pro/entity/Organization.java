package com.xiaofeng.pro.entity;

import com.xiaofeng.pro.base.MyBaseEntity;
import com.xiaofeng.pro.common.annotation.ExcelField;
import lombok.Data;
import lombok.EqualsAndHashCode;

import javax.persistence.*;
import javax.validation.constraints.Email;
import javax.validation.constraints.Pattern;

/**
 * @program: people
 * @description: 机构
 * @author: dongxiaofeng
 * @create: 2019-01-07-11-53
 */
@Entity
@Data
@Table(name = "t_organization")
public class Organization extends MyBaseEntity {


    @Column(unique = true)
    @ExcelField(title = "机构名称", align = 0, sort = 2)
    private String name;


    @Column(unique = true)
    @ExcelField(title = "编码", align = 1, sort = 0)
    private String code;

    //修改的时候不更新
    @Column(updatable = false)
    @ExcelField(title = "机构id", align = 2, sort = 1)
    private Long organizationTypeId;

    private Long parentId;

    private String  address;

    private String postalcode;

    @Pattern(regexp = "1[3|4|5|7|8][0-9]\\d{8}", message = "手机号格式错误")
    private String telephone;

    @Email
    private String email;

    private String contactName;

    private String contactPosition;

    private Integer status;

    //机构拥有的默认超级管理员，在添加机构的时候有用自动匹配属性,其他列表啥的暂时不用先
    @Transient
    private User user;

    @Transient
    @ManyToOne
    @JoinColumn(name = "organization_type_id", referencedColumnName = "id")
    private OrganizationType organizationType;
}
