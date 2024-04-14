# -*- coding: utf-8 -*-
"""
Created on Sat Sep 29 10:17:40 2018

@author: gonhjian
"""

import win32com
import win32com.client
#import re
#import matplotlib
#import numpy as np
#import matplotlib.cm as cm
#import matplotlib.mlab as mlab
#import matplotlib.pyplot as plt
#import pandas as pd
#from mpl_toolkits.mplot3d import Axes3D
#from matplotlib.mlab import griddata
#from matplotlib.colors import Normalize
#from scipy.interpolate import griddata as gd
#from scipy import interpolate 
import math
import os
#import sys

 

        
def Contour(datafile,outfile,blank_file,xcol,ycol,zcol,algorithm,showreport,xlength,ylength,top,left,caxshowlabel1,caxshowlabel2,baxshowlabel1,baxshowlabel2,caxmajorticklength1,\
            caxmajorticklength2,baxmajorticklength1,baxmajorticklength2,caxlinestyle1,caxlinestyle2,baxlinestyle1,baxlinestyle2,caxwidth1,caxwidth2,caxmajorticktype1,caxmajorticktype2,\
            baxmajorticktype1,baxmajorticktype2,caxbtitle,caxltitle,caxbtitlefontsize,caxltitlefontsize,caxbtitlefontface,caxltitlefontface,caxblabelfontsize,caxllabelfontsize,caxblabelface,\
            caxllabelface,caxbtitleoffset2,caxltitleoffset2,caxbtitleoffset1,caxltitleoffset1,caxbtitlefontbold,caxltitlefontbold,caxblabelfontbold,caxllabelfontbold,\
            clfillcontour,clsmoothcontour,interval,cllevelmajorinterval,clshowmajorlabels,clshowminorlabels,clshowcolorscale,clorientlabelsuphill,clopacity,clminlinestyle,clmajlinestyle,\
            clmajlinewidth,cllabelfontface,cllabelfontsize,cllabelfontbold,save_file,baselinewith,numcols,numrows,colormappath):
    '''
    excel_file          输入边界文件.xlsx                       bln_file                输出边界文件.bln
    bna_file            输出边界文件.bna                        datafile                输入原始文件.dat
    outfile             输出网格文件.grd                        blank_file              输入白化文件.bln,值为bln_file
    xcol                输入x轴值                               ycol                    输入y轴值
    zcol                输入z轴值                               algorithm               输入插值算法
    showreport          是否显示网格文件                         xlength                 输入图框长
    ylength             输入图框高                              top                     输入图框位置，上边距
    left                输入图框位置,左边距                      caxshowlabel1           输入图框左侧和底侧坐标轴是否显示标签
    caxshowlabel2       输入图框上侧和右侧坐标轴是否显示标签      baxshowlabel1           输入底图左侧和底侧坐标轴是否显示标签
    baxshowlabel2       输入底图上侧和右侧坐标轴是否显示标签      caxmajorticklength1     输入图框左侧和底侧坐标轴的主刻线长度
    caxmajorticklength2 输入图框上侧和右侧坐标轴的主刻线长度      baxmajorticklength1     输入底图左侧和底侧坐标轴的主刻线长度
    baxmajorticklength2 输入底图上侧和右侧坐标轴的主刻线长度      caxlinestyle1           输入图框左侧和底侧坐标轴线线型
    caxlinestyle2       输入图框上侧和右侧坐标轴线线型            baxlinestyle1           输入底图左侧和底侧坐标轴线线型
    baxlinestyle2       输入底图上侧和右侧坐标轴线线型            caxwidth1               输入图框左侧和底侧坐标轴线宽
    caxwidth2           输入图框上侧和右侧坐标轴线宽              caxmajorticktype1       输入图框左侧和底侧坐标轴刻线样式
    caxmajorticktype2   输入图框上侧和右侧坐标轴刻线样式          baxmajorticktype1       输入底图左侧和底侧坐标轴刻线样式
    baxmajorticktype2   输入底图上侧和右侧坐标轴刻线样式          caxbtitle               输入底部坐标轴名称
    caxltitle           输入左侧坐标轴名称                       caxbtitlefontsize       输入底侧坐标轴名称字体大小
    caxltitlefontsize   输入左侧坐标轴名称字体大小               caxbtitlefontface        输入底侧坐标轴名称字体样式
    caxltitlefontface   输入左侧坐标轴名称字体样式               caxblabelfontsize        输入底侧坐标轴标签字体大小
    caxllabelfontsize   输入左侧坐标轴标签字体大小               caxblabelface            输入底侧坐标轴标签字体样式
    caxllabelface       输入左侧坐标轴标签字体样式               caxbtitleoffset2         输入底侧坐标轴名称上下移动距离
    caxltitleoffset2    输入左侧坐标轴名称上下移动距离           caxbtitleoffset1         输入底侧坐标轴名称左右移动距离
    caxltitleoffset1    输入左侧坐标轴名称左右移动距离           caxbtitlefontbold        输入底侧坐标轴名称是否加粗
    caxltitlefontbold   输入左侧坐标轴名称是否加粗               caxblabelfontbold        输入底侧坐标轴标签是否加粗
    caxllabelfontbold   输入左侧坐标轴标签是否加粗               clfillcontour            输入等高线图层是否着色
    clsmoothcontour     输入等高线平滑程度                      interval                 输入等高线最小间距
    cllevelmajorinterval输入主等高线的间隔                      clshowmajorlabels        输入主等高线是否显示标签
    clshowminorlabels   输入次等高线是否显示标签                clshowcolorscale          输入是否显示颜色标签
    clorientlabelsuphill输入是否倒置等高线标签                  clopacity                 输入等高线图层透明度
    clminlinestyle      输入次等高线线型                        clmajlinestyle           输入主等高线线型 
    clmajlinewidth      输入主等高线线宽                        cllabelfontface          输入等高线标签字体样式
    cllabelfontsize     输入等高线标签字体大小                  cllabelfontbold           输入等高线标签字体是否加粗
    save_file           保存文件.srf                           baselinewith              输入轮廓线宽
    colormappath        读取云图色阶文件.clr
    '''
    
    app =win32com.client.gencache.EnsureDispatch("Surfer.Application")
    Plot=app.Documents.Add(1)
    app.Visible = True
    DataFile = datafile
    
    OutFile =os.path.abspath(outfile)
    blankfile=blank_file
    savefile=save_file
    
    if datafile!=0:        
        Grid =app.NewGrid()
        app.GridData (DataFile=DataFile,xCol=xcol,yCol=ycol,zCol=zcol,NumCols=numcols,NumRows=numrows,Algorithm =algorithm,ShowReport=showreport, OutGrid= OutFile)
        app.GridBlank(InGrid=OutFile,BlankFile=blankfile,OutGrid=OutFile,OutFmt=3)
#       MapFrame1=Plot.Shapes.AddBaseMap(blankfile)
        Grid.LoadFile(OutFile,HeaderOnly=True)
        #Creates a contour map and assigns the map frame to the variable "MapFrame"
        MapFrame = Plot.Shapes.AddContourMap(OutFile)
#        Plot.Shapes.SelectAll() 
#        Plot.Selection.OverlayMaps()
        ContourMap=MapFrame.Overlays(1)
#       Plot.Shapes.SelectAll() 
#       Plot.Selection.OverlayMaps()
        #检索出等值线图，修改其属性
        x1=abs(Grid.xMax)
        x2=abs(Grid.xMin)
        y1=abs(Grid.yMax)
        y2=abs(Grid.yMin)
        z1=abs(Grid.zMax)
        z2=abs(Grid.zMin)
        if (x1>=1):
            x1=x1
            xmax=math.ceil(Grid.xMax)
            # xmax=230
        else:
            i1=0
            while (x1<1):
                i1=i1+1
                x1=x1*math.pow(10,i1)
            xmax=math.ceil(Grid.xMax*math.pow(10,i1))/(1.0*math.pow(10,i1))
        if (x2>=1):
            x2=x2
            xmin=math.floor(Grid.xMin)
            # xmin=0
        else:
            i2=0
            while (x2<1):
                i2=i2+1
                x2=x2*math.pow(10,i2)
            xmin=math.floor(Grid.xMin*math.pow(10,i2))/(1.0*math.pow(10,i2))
            
        if (y1>=1):
            y1=y1
            ymax=math.ceil(Grid.yMax)
        else:
            j1=0
            while (x1<1):
                j1=j1+1
                y1=y1*math.pow(10,j1)
            ymax=math.ceil(Grid.yMax*pow(10,j1))/(1.0*math.pow(10,j1))
        if (y2>=1):
            y2=y2
            # ymin=math.floor(Grid.yMin)
            ymin=0
        else:
            j2=0
            while (y2<1):
                j2=j2+1
                y2=y2*math.pow(10,j2)
            ymin=math.floor(Grid.yMin*math.pow(10,j2))/(1.0*math.pow(10,j2))
        if (z1>=1):
            z1=z1
            zmax=math.ceil(Grid.zMax)
        else:
            k1=0
            while (z1<1):
                k1=k1+1
                z1=z1*math.pow(10,k1)
            zmax=math.ceil(Grid.zMax*math.pow(10,k1))/(1.0*math.pow(10,k1))
        if (z2>=1):
            z2=z2
            zmin=math.floor(Grid.zMin)
        else:
            k2=0
            while (z2<1):
                k2=k2+1
                z2=z2*math.pow(10,k2) 
            zmin=math.floor(Grid.zMin*math.pow(10,k2))/(1.0*math.pow(10,k2))

    
        #Changes the Mapframe Property
        MapFrame.SetLimits (xMin=xmin, xMax=xmax, yMin=ymin, yMax=ymax)
        MapFrame.xLength=xlength
        MapFrame.yLength=ylength
        MapFrame.Top=top
        MapFrame.Left =left
        for Axis in ['Bottom Axis','Left Axis']:  #,'Top Axis','Right Axis']:
            MapFrame.Axes(Axis).ShowLabels =caxshowlabel1
            MapFrame.Axes(Axis).MajorTickLength=caxmajorticklength1
            MapFrame.Axes(Axis).AxisLine.Style =caxlinestyle1
            MapFrame.Axes(Axis).AxisLine.Width =caxwidth1
            MapFrame.Axes(Axis).MajorTickType =caxmajorticktype1
        for Axis in ['Top Axis','Right Axis']:
            MapFrame.Axes(Axis).ShowLabels =caxshowlabel2
            MapFrame.Axes(Axis).MajorTickType =caxmajorticktype2 
            MapFrame.Axes(Axis).AxisLine.Style =caxlinestyle2
            MapFrame.Axes(Axis).MajorTickLength=caxmajorticklength2
            MapFrame.Axes(Axis).AxisLine.Width =caxwidth2
        MapFrame.Axes("Bottom Axis").Title = caxbtitle
        MapFrame.Axes("Left Axis").Title= caxltitle
        MapFrame.Axes("Bottom Axis").TitleFont.Size =caxbtitlefontsize
        MapFrame.Axes("Left Axis").TitleFont.Size= caxltitlefontsize
        MapFrame.Axes("Bottom Axis").TitleFont.Face =caxbtitlefontface
        MapFrame.Axes("Left Axis").TitleFont.Face=caxltitlefontface
        MapFrame.Axes("Bottom Axis").LabelFont.Size =caxblabelfontsize
        MapFrame.Axes("Left Axis").LabelFont.Size= caxllabelfontsize    
        MapFrame.Axes("Bottom Axis").LabelFont.Face =caxblabelface
        MapFrame.Axes("Left Axis").LabelFont.Face=caxllabelface
        MapFrame.Axes("Bottom Axis").TitleOffset2=caxbtitleoffset2
        MapFrame.Axes("Left Axis").TitleOffset2=caxltitleoffset2
        MapFrame.Axes("Bottom Axis").TitleOffset1=caxbtitleoffset1
        MapFrame.Axes("Left Axis").TitleOffset1=caxltitleoffset1
        MapFrame.Axes("Bottom Axis").TitleFont.Bold =caxbtitlefontbold
        MapFrame.Axes("Left Axis").TitleFont.Bold= caxltitlefontbold
        MapFrame.Axes("Bottom Axis").LabelFont.Bold =caxblabelfontbold
        MapFrame.Axes("Left Axis").LabelFont.Bold= caxllabelfontbold
 
        
        ###将ContourLayer属性赋给ContourMap
        ContourLayer = win32com.client.CastTo(ContourMap, "IContourLayer")
    
        #将Changes the ContourLayer Property
        ContourLayer.FillContours=clfillcontour   
        ContourLayer.FillForegroundColorMap.LoadFile(colormappath)   
        ContourLayer.SmoothContours =clsmoothcontour
        ContourLayer.SetSimpleLevels(Min=zmin, Max=zmax, Interval=interval)    
        #设置等高线的level
        ContourLayer.LevelMajorInterval=cllevelmajorinterval                                  
        #设置主等高线的间隔、
        ContourLayer.ShowMajorLabels =clshowmajorlabels                                
        #设置主等高线的标签的开关
        ContourLayer.ShowMinorLabels =clshowminorlabels                               
        #设置次等高线的标签的开关
        ContourLayer.ShowColorScale =clshowcolorscale                                
        #设置层次标签
        ContourLayer.OrientLabelsUphill =clorientlabelsuphill                            
        #倒置等高线标签
        ContourLayer.Opacity =clopacity                                         
        #等高线透明度
        ContourLayer.MinorLine.Style=clminlinestyle                           
        #次等高线线型
        ContourLayer.MajorLine.Style=clmajlinestyle                               
        #主等高线线型
        ContourLayer.MajorLine.Width=clmajlinewidth                                     
        #主等高线线宽
        ContourLayer.LabelFont.Face =cllabelfontface                   
        #主等高线标签字体
        ContourLayer.LabelFont.Size=cllabelfontsize                                    
        #等高线标签字体大小
        ContourLayer.LabelFont.Bold=cllabelfontbold                                    
        #等高线标签字体加粗
#       ContourLayer.LabelTolerance =2                                    
        #标签曲线公差
#       ContourLayer.LabelLabelDist = 1.25                                 
        #标签间距
#       ContourLayer.LabelEdgeDist = 0.75                                
        #标签与图边缘间距 
#       ContourLayer.HachLength = 1                                     
        #刻线长度
#       ContourLayer.HachDirection=2                                      
        #刻线方向
#       MapFrame=Plot.Shapes.AddBaseLayer(Map= MapFrame, ImportFileName=blankfile)
#       BaseLayer = win32com.client.CastTo(BaseMap, "IBaseLayer")
        MapFrame=Plot.Shapes.AddBaseMap(blankfile)
        MapFrame.SetLimits (xMin=xmin, xMax=xmax, yMin=ymin, yMax=ymax)
        MapFrame.xLength=xlength
        MapFrame.yLength=ylength
        MapFrame.Top=top
        MapFrame.Left =left
        for Axis in ['Bottom Axis','Left Axis']:  #,'Top Axis','Right Axis']:
            MapFrame.Axes(Axis).ShowLabels =baxshowlabel1
            MapFrame.Axes(Axis).MajorTickLength=baxmajorticklength1
            MapFrame.Axes(Axis).MajorTickType =baxmajorticktype1
            MapFrame.Axes(Axis).AxisLine.Style =baxlinestyle1
#           MapFrame.Axes(Axis).MajorTickLength=0.05
        for Axis in ['Top Axis','Right Axis']:
            MapFrame.Axes(Axis).MajorTickType =baxmajorticktype2 
            MapFrame.Axes(Axis).ShowLabels =baxshowlabel2
            MapFrame.Axes(Axis).MajorTickLength=baxmajorticklength2
            MapFrame.Axes(Axis).AxisLine.Style =baxlinestyle2
            Plot.SaveAs(FileName=savefile, FileFormat=17)
        BaseMap=MapFrame.Overlays(1)
        Baselayer = win32com.client.CastTo(BaseMap, "IBaseLayer")
        Baselayer.Line.Width=baselinewith
    else:
        Grid =app.NewGrid()
#       app.GridData (DataFile=DataFile,xCol=xcol,yCol=ycol,zCol=zcol,Algorithm =algorithm,ShowReport=showreport, OutGrid= OutFile)
        app.GridBlank(InGrid=OutFile,BlankFile=blankfile,OutGrid=OutFile,OutFmt=3)
#       MapFrame1=Plot.Shapes.AddBaseMap(blankfile)
        Grid.LoadFile(OutFile,HeaderOnly=True)
        #Creates a contour map and assigns the map frame to the variable "MapFrame"
        MapFrame = Plot.Shapes.AddContourMap(OutFile)
#        Plot.Shapes.SelectAll() 
#        Plot.Selection.OverlayMaps()
        ContourMap=MapFrame.Overlays(1)
#       Plot.Shapes.SelectAll() 
#       Plot.Selection.OverlayMaps()
        #检索出等值线图，修改其属性
        x1=abs(Grid.xMax)
        x2=abs(Grid.xMin)
        y1=abs(Grid.yMax)
        y2=abs(Grid.yMin)
        z1=abs(Grid.zMax)
        z2=abs(Grid.zMin)
        if (x1>=1):
            x1=x1
            xmax=math.ceil(Grid.xMax)
        else:
            i1=0
            while (x1<1):
                i1=i1+1
                x1=x1*math.pow(10,i1)
            xmax=math.ceil(Grid.xMax*math.pow(10,i1))/(1.0*math.pow(10,i1))
        if (x2>=1):
            x2=x2
            xmin=math.floor(Grid.xMin)
        else:
            i2=0
            while (x2<1):
                i2=i2+1
                x2=x2*math.pow(10,i2)
            xmin=math.floor(Grid.xMin*math.pow(10,i2))/(1.0*math.pow(10,i2))
            
        if (y1>=1):
            y1=y1
            ymax=math.ceil(Grid.yMax)
        else:
            j1=0
            while (x1<1):
                j1=j1+1
                y1=y1*math.pow(10,j1)
            ymax=math.ceil(Grid.yMax*pow(10,j1))/(1.0*math.pow(10,j1))
        if (y2>=1):
            y2=y2
            ymin=math.floor(Grid.yMin)
        else:
            j2=0
            while (y2<1):
                j2=j2+1
                y2=y2*math.pow(10,j2)
            ymin=math.floor(Grid.yMin*math.pow(10,j2))/(1.0*math.pow(10,j2))
        if (z1>=1):
            z1=z1
            zmax=math.ceil(Grid.zMax)
        else:
            k1=0
            while (z1<1):
                k1=k1+1
                z1=z1*math.pow(10,k1)
            zmax=math.ceil(Grid.zMax*math.pow(10,k1))/(1.0*math.pow(10,k1))
        if (z2>=1):
            z2=z2
            zmin=math.floor(Grid.zMin)
        else:
            k2=0
            while (z2<1):
                k2=k2+1
                z2=z2*math.pow(10,k2) 
            zmin=math.floor(Grid.zMin*math.pow(10,k2))/(1.0*math.pow(10,k2))
    
        #Changes the Mapframe Property
        MapFrame.SetLimits (xMin=xmin, xMax=xmax, yMin=ymin, yMax=ymax)
        MapFrame.xLength=xlength
        MapFrame.yLength=ylength
        MapFrame.Top=top
        MapFrame.Left =left
        for Axis in ['Bottom Axis','Left Axis']:  #,'Top Axis','Right Axis']:
            MapFrame.Axes(Axis).ShowLabels =caxshowlabel1
            MapFrame.Axes(Axis).MajorTickLength=caxmajorticklength1
            MapFrame.Axes(Axis).AxisLine.Style =caxlinestyle1
            MapFrame.Axes(Axis).AxisLine.Width =caxwidth1
            MapFrame.Axes(Axis).MajorTickType =caxmajorticktype1
        for Axis in ['Top Axis','Right Axis']:
            MapFrame.Axes(Axis).ShowLabels =caxshowlabel2
            MapFrame.Axes(Axis).MajorTickType =caxmajorticktype2 
            MapFrame.Axes(Axis).AxisLine.Style =caxlinestyle2
            MapFrame.Axes(Axis).MajorTickLength=caxmajorticklength2
            MapFrame.Axes(Axis).AxisLine.Width =caxwidth2
        MapFrame.Axes("Bottom Axis").Title = caxbtitle
        MapFrame.Axes("Left Axis").Title= caxltitle
        MapFrame.Axes("Bottom Axis").TitleFont.Size =caxbtitlefontsize
        MapFrame.Axes("Left Axis").TitleFont.Size= caxltitlefontsize
        MapFrame.Axes("Bottom Axis").TitleFont.Face =caxbtitlefontface
        MapFrame.Axes("Left Axis").TitleFont.Face=caxltitlefontface
        MapFrame.Axes("Bottom Axis").LabelFont.Size =caxblabelfontsize
        MapFrame.Axes("Left Axis").LabelFont.Size= caxllabelfontsize    
        MapFrame.Axes("Bottom Axis").LabelFont.Face =caxblabelface
        MapFrame.Axes("Left Axis").LabelFont.Face=caxllabelface
        MapFrame.Axes("Bottom Axis").TitleOffset2=caxbtitleoffset2
        MapFrame.Axes("Left Axis").TitleOffset2=caxltitleoffset2
        MapFrame.Axes("Bottom Axis").TitleOffset1=caxbtitleoffset1
        MapFrame.Axes("Left Axis").TitleOffset1=caxltitleoffset1
        MapFrame.Axes("Bottom Axis").TitleFont.Bold =caxbtitlefontbold
        MapFrame.Axes("Left Axis").TitleFont.Bold= caxltitlefontbold
        MapFrame.Axes("Bottom Axis").LabelFont.Bold =caxblabelfontbold
        MapFrame.Axes("Left Axis").LabelFont.Bold= caxllabelfontbold 
        
        ###将ContourLayer属性赋给ContourMap
        ContourLayer = win32com.client.CastTo(ContourMap, "IContourLayer")
    
        #将Changes the ContourLayer Property
        ContourLayer.FillContours=clfillcontour    
        ContourLayer.FillForegroundColorMap.LoadFile(colormappath)    
        ContourLayer.SmoothContours =clsmoothcontour
        ContourLayer.SetSimpleLevels(Min=zmin, Max=zmax, Interval=interval)    
        #设置等高线的level
        ContourLayer.LevelMajorInterval=cllevelmajorinterval                                  
        #设置主等高线的间隔
        ContourLayer.ShowMajorLabels =clshowmajorlabels                                
        #设置主等高线的标签的开关
        ContourLayer.ShowMinorLabels =clshowminorlabels                               
        #设置次等高线的标签的开关
        ContourLayer.ShowColorScale =clshowcolorscale                                
        #设置层次标签
        ContourLayer.OrientLabelsUphill =clorientlabelsuphill                            
        #倒置等高线标签
        ContourLayer.Opacity =clopacity                                         
        #等高线透明度
        ContourLayer.MinorLine.Style=clminlinestyle                           
        #次等高线线型
        ContourLayer.MajorLine.Style=clmajlinestyle                               
        #主等高线线型
        ContourLayer.MajorLine.Width=clmajlinewidth                                     
        #主等高线线宽
        ContourLayer.LabelFont.Face =cllabelfontface                   
        #主等高线标签字体
        ContourLayer.LabelFont.Size=cllabelfontsize                                    
        #等高线标签字体大小
        ContourLayer.LabelFont.Bold=cllabelfontbold                                    
        #等高线标签字体加粗
#       ContourLayer.LabelTolerance =2                                    
        #标签曲线公差
#       ContourLayer.LabelLabelDist = 1.25                                 
        #标签间距
#       ContourLayer.LabelEdgeDist = 0.75                                
        #标签与图边缘间距 
#       ContourLayer.HachLength = 1                                     
        #刻线长度
#       ContourLayer.HachDirection=2                                      
        #刻线方向
#       MapFrame=Plot.Shapes.AddBaseLayer(Map= MapFrame, ImportFileName=blankfile)
#       BaseLayer = win32com.client.CastTo(BaseMap, "IBaseLayer")
        MapFrame=Plot.Shapes.AddBaseMap(blankfile)
        MapFrame.SetLimits (xMin=xmin, xMax=xmax, yMin=ymin, yMax=ymax)
        MapFrame.xLength=xlength
        MapFrame.yLength=ylength
        MapFrame.Top=top
        MapFrame.Left =left
        for Axis in ['Bottom Axis','Left Axis']:  #,'Top Axis','Right Axis']:
            MapFrame.Axes(Axis).ShowLabels =baxshowlabel1
            MapFrame.Axes(Axis).MajorTickLength=baxmajorticklength1
            MapFrame.Axes(Axis).MajorTickType =baxmajorticktype1
            MapFrame.Axes(Axis).AxisLine.Style =baxlinestyle1
#           MapFrame.Axes(Axis).MajorTickLength=0.05
        for Axis in ['Top Axis','Right Axis']:
            MapFrame.Axes(Axis).MajorTickType =baxmajorticktype2 
            MapFrame.Axes(Axis).ShowLabels =baxshowlabel2
            MapFrame.Axes(Axis).MajorTickLength=baxmajorticklength2
            MapFrame.Axes(Axis).AxisLine.Style =baxlinestyle2
            Plot.SaveAs(FileName=savefile, FileFormat=17)
        BaseMap=MapFrame.Overlays(1)
        Baselayer = win32com.client.CastTo(BaseMap, "IBaseLayer")
        Baselayer.Line.Width=baselinewith


def BLN(excel_file,bln_file,bna_file):
    app =win32com.client.gencache.EnsureDispatch("Surfer.Application")
    Plot=app.Documents.Open2(excel_file,Options="Sheet=Sheet1", FilterId="xlsx")
    app.Visible = True
    Plot.SaveAs(FileName=bln_file, FileFormat=12)
    Plot.SaveAs(FileName=bna_file, FileFormat=14)
    Plot.Close(1)
'''
    excel_file          输入边界文件.xlsx                       bln_file                输出边界文件.bln
    bna_file            输出边界文件.bna                        datafile                输入原始文件.dat
    outfile             输出网格文件.grd                        blank_file              输入白化文件.bln,值为bln_file
    xcol                输入x轴值                               ycol                    输入y轴值
    zcol                输入z轴值                               algorithm               输入插值算法
    showreport          是否显示网格文件                         xlength                 输入图框长
    ylength             输入图框高                              top                     输入图框位置，上边距
    left                输入图框位置,左边距                      caxshowlabel1           输入图框左侧和底侧坐标轴是否显示标签
    caxshowlabel2       输入图框上侧和右侧坐标轴是否显示标签      baxshowlabel1           输入底图左侧和底侧坐标轴是否显示标签
    baxshowlabel2       输入底图上侧和右侧坐标轴是否显示标签      caxmajorticklength1     输入图框左侧和底侧坐标轴的主刻线长度
    caxmajorticklength2 输入图框上侧和右侧坐标轴的主刻线长度      baxmajorticklength1     输入底图左侧和底侧坐标轴的主刻线长度
    baxmajorticklength2 输入底图上侧和右侧坐标轴的主刻线长度      caxlinestyle1           输入图框左侧和底侧坐标轴线线型
    caxlinestyle2       输入图框上侧和右侧坐标轴线线型            baxlinestyle1           输入底图左侧和底侧坐标轴线线型
    baxlinestyle2       输入底图上侧和右侧坐标轴线线型            caxwidth1               输入图框左侧和底侧坐标轴线宽
    caxwidth2           输入图框上侧和右侧坐标轴线宽              caxmajorticktype1       输入图框左侧和底侧坐标轴刻线样式
    caxmajorticktype2   输入图框上侧和右侧坐标轴刻线样式          baxmajorticktype1       输入底图左侧和底侧坐标轴刻线样式
    baxmajorticktype2   输入底图上侧和右侧坐标轴刻线样式          caxbtitle               输入底部坐标轴名称
    caxltitle           输入左侧坐标轴名称                       caxbtitlefontsize       输入底侧坐标轴名称字体大小
    caxltitlefontsize   输入左侧坐标轴名称字体大小               caxbtitlefontface        输入底侧坐标轴名称字体样式
    caxltitlefontface   输入左侧坐标轴名称字体样式               caxblabelfontsize        输入底侧坐标轴标签字体大小
    caxllabelfontsize   输入左侧坐标轴标签字体大小               caxblabelface            输入底侧坐标轴标签字体样式
    caxllabelface       输入左侧坐标轴标签字体样式               caxbtitleoffset2         输入底侧坐标轴名称上下移动距离
    caxltitleoffset2    输入左侧坐标轴名称上下移动距离           caxbtitleoffset1         输入底侧坐标轴名称左右移动距离
    caxltitleoffset1    输入左侧坐标轴名称左右移动距离           caxbtitlefontbold        输入底侧坐标轴名称是否加粗
    caxltitlefontbold   输入左侧坐标轴名称是否加粗               caxblabelfontbold        输入底侧坐标轴标签是否加粗
    caxllabelfontbold   输入左侧坐标轴标签是否加粗               clfillcontour            输入等高线图层是否着色
    clsmoothcontour     输入等高线平滑程度                      interval                 输入等高线最小间距
    cllevelmajorinterval输入主等高线的间隔                      clshowmajorlabels        输入主等高线是否显示标签
    clshowminorlabels   输入次等高线是否显示标签                clshowcolorscale          输入是否显示颜色标签
    clorientlabelsuphill输入是否倒置等高线标签                  clopacity                 输入等高线图层透明度
    clminlinestyle      输入次等高线线型                        clmajlinestyle           输入主等高线线型 
    clmajlinewidth      输入主等高线线宽                        cllabelfontface          输入等高线标签字体样式
    cllabelfontsize     输入等高线标签字体大小                  cllabelfontbold           输入等高线标签字体是否加粗
    save_file           保存文件.srf                           baselinewith              输入轮廓线宽
'''    
    
    

BLN(excel_file=r"C:\Users\Gj\Desktop\excel1.xlsx",bln_file=r"C:\Users\Gj\Desktop\sec12.bln",bna_file=r"C:\Users\Gj\Desktop\sec12.bna")

Contour(datafile=r"C:\Users\Gj\Desktop\CALCULATIONSTAGES KK=  15THE PRINCIPAL STRESSES OF THE ROCKFILLSECTION NO.= 5ELEMENT NO.  X-COORDI  Y-COORDI  Z-COORDICGM1CGM2CGM3CGMXCGMYCGMZSTRESS LEVEL.dat",outfile= r"C:\Users\Gj\Desktop\sec12.grd",\
        blank_file=r"C:\Users\Gj\Desktop\sec12.bln",xcol=2,ycol=3,zcol=5,algorithm=9,showreport=False,\
        xlength=20,ylength=4,top=7.25,left=-1.1,caxshowlabel1=True,caxshowlabel2=False,baxshowlabel1=False,\
        baxshowlabel2=False,caxmajorticklength1=0.05,caxmajorticklength2=0,baxmajorticklength1=0,baxmajorticklength2=0,\
        caxlinestyle1='Solid',caxlinestyle2='Solid',baxlinestyle1='Invisible',baxlinestyle2='Invisible',caxwidth1=0,\
        caxwidth2=0,caxmajorticktype1=2,caxmajorticktype2=1,baxmajorticktype1=1,baxmajorticktype2=2,caxbtitle='x(m)',\
        caxltitle='y(m)',caxbtitlefontsize=15,caxltitlefontsize=15,caxbtitlefontface="Times New Roman",caxltitlefontface="Times New Roman",\
        caxblabelfontsize=13,caxllabelfontsize=13,caxblabelface="Times New Roman",caxllabelface="Times New Roman",caxbtitleoffset2=-0.05,\
        caxltitleoffset2=0,caxbtitleoffset1=0,caxltitleoffset1=0,caxbtitlefontbold=False, caxltitlefontbold=False,caxblabelfontbold=False,\
        caxllabelfontbold=False,clfillcontour=True,clsmoothcontour=4,interval=0.1,cllevelmajorinterval=1,\
        clshowmajorlabels=True,clshowminorlabels=False,clshowcolorscale=False,clorientlabelsuphill=False,clopacity=60,clminlinestyle='Invisible',\
        clmajlinestyle='Solid',clmajlinewidth=0,cllabelfontface="Times New Roman",cllabelfontsize=10,cllabelfontbold=False,\
        save_file=r"C:\Users\Gj\Desktop\sec12.srf",baselinewith=0.04,numcols=1000,numrows=1000,colormappath=r"C:\Program Files\Golden Software\Surfer 14\ColorScales\Rainbow3.clr")



'''
Contour(datafile=0,outfile= "C:\\Users\gonhjian\Desktop\out11.grd",\
        blank_file="C:\\Users\gonhjian\Desktop\sec418.bln",xcol=2,ycol=3,zcol=5,algorithm=9,showreport=False,\
        xlength=10,ylength=4,top=7.25,left=-1.1,caxshowlabel1=True,caxshowlabel2=False,baxshowlabel1=False,\
        baxshowlabel2=False,caxmajorticklength1=0.05,caxmajorticklength2=0,baxmajorticklength1=0,baxmajorticklength2=0,\
        caxlinestyle1='Solid',caxlinestyle2='Solid',baxlinestyle1='Invisible',baxlinestyle2='Invisible',caxwidth1=0,\
        caxwidth2=0,caxmajorticktype1=2,caxmajorticktype2=1,baxmajorticktype1=1,baxmajorticktype2=2,caxbtitle='x(m)',\
        caxltitle='y(m)',caxbtitlefontsize=15,caxltitlefontsize=15,caxbtitlefontface="Times New Roman",caxltitlefontface="Times New Roman",\
        caxblabelfontsize=13,caxllabelfontsize=13,caxblabelface="Times New Roman",caxllabelface="Times New Roman",caxbtitleoffset2=-0.05,\
        caxltitleoffset2=0,caxbtitleoffset1=0,caxltitleoffset1=0,caxbtitlefontbold=False, caxltitlefontbold=False,caxblabelfontbold=False,\
        caxllabelfontbold=False,caxaxislinewidth=0,clfillcontour=False,clsmoothcontour=4,interval=0.7,cllevelmajorinterval=1,\
        clshowmajorlabels=True,clshowminorlabels=False,clshowcolorscale=False,clorientlabelsuphill=False,clopacity=60,clminlinestyle='Invisible',\
        clmajlinestyle='Solid',clmajlinewidth=0,cllabelfontface="Times New Roman",cllabelfontsize=10,cllabelfontbold=False,\
        save_file="C:\\Users\gonhjian\Desktop\kk.srf")
'''    