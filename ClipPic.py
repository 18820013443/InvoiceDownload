import cv2 as cv
import numpy as np
import os

class Pic:
    # def __init__(self,strInputImgPath, strOutputImgPath):
    def __init__(self):
        # self.strInputImgPath = strInputImgPath
        # self.strOutputImgPath = strOutputImgPath
        pass

    def ClipImg(self, strInputImgPath, strOutputImgPath):
        # strImgPath = os.path.join(os.getcwd(), r'ExamplePics\example.JPG')
        # img = cv.imread(strImgPath)
        img = cv.imread(strInputImgPath)
        imgGs = cv.GaussianBlur(img, (3,3), 0)
        imgGray = cv.cvtColor(imgGs, cv.COLOR_BGR2GRAY)
        th = cv.adaptiveThreshold(imgGray, 255, cv.ADAPTIVE_THRESH_GAUSSIAN_C, cv.THRESH_BINARY, 11, 4)
        edges = cv.Canny(th, 10, 200)
        contours, hierarchy = cv.findContours(edges, cv.RETR_CCOMP, cv.CHAIN_APPROX_SIMPLE)

        rectList, rectAreaList = [], []
        for i in range(len(contours)):
            rect = cv.boundingRect(contours[i])
            w = rect[2]
            h = rect[3]
            area = w * h
            rectAreaList.append(area)
            rectList.append(rect)
        areaMax = max(rectAreaList)
        post = rectAreaList.index(areaMax)
        x = rectList[post][0]
        y = rectList[post][1]
        w = rectList[post][0] + rectList[post][2]
        h = rectList[post][1] + rectList[post][3]

        clipImg = img[y:h, x:w]
        # cv.imwrite('result.jpg', clipImg)
        cv.imwrite(strOutputImgPath, clipImg)
        # cv.imshow('result', clipImg)
        # cv.waitKey(0)
        # cv.destroyAllwindows()


if __name__ == '__main__':
    pass
    # strInputImgPath = os.path.join(os.getcwd(), r'ExamplePics\example.JPG')
    # strOutputImgPath = os.path.join(os.getcwd(), r'ExamplePics\result.jpg')
    # objPic = Pic()
    # objPic.ClipImg(strInputImgPath, strOutputImgPath)


