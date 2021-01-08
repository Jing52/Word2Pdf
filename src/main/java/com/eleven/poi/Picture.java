package com.eleven.poi;

import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;
import lombok.ToString;

/**
 * @Author: Evan
 * @CreateTime: 2021-01-08
 * @Description:
 */
@Setter
@Getter
@ToString
@NoArgsConstructor
public class Picture {
    private String image;
    private String pictureType;
    private String width;
    private String height;

    public Picture(String url, String type) {
        this.image = url;
        this.pictureType = type;
    }
}
