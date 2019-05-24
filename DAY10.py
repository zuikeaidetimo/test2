from tkinter import *
from tkinter.filedialog import *
import os.path
import math

###########################################################
########### RAW 데이터 분석#####################
###########################################################
## 함수 선언부
def loadImage(fname) :
    global window, canvas, paper, inW, inH, outW, outH, inImage, outImage
    inImage = [] # 초기화
    # 파일 크기 계산
    fsize = os.path.getsize(fname) # Byte 단위
    inW = inH = int(math.sqrt(fsize))
    # 빈 메모리 확보 (2차원 리스트)
    for _ in range(inH) :
        tmp = []
        for _ in range(inW) :
            tmp.append(0)
        inImage.append(tmp)
    # 파일 --> 메모리로 한개씩 옮기기
    fp = open(fname, 'rb')
    for  i  in  range(inH) :
        for k in range(inW) :
            data = int(ord(fp.read(1))) # 1개 픽셀값을 읽음 (0~255)
            inImage[i][k] = data
    fp.close()


def  openImage() :
    global window, canvas, paper, inW, inH, outW, outH, inImage, outImage, filename
    filename = askopenfilename(parent=window, filetypes=(("RAW 파일", "*.raw"), ("모든 파일", "*.*")))
    if filename == "" or filename == None :
        return


    # 파일 --> 메모리
    loadImage(filename)

    # Input --> outPut으로 동일하게 만들기.
    equalImage()

def displayImage() :
    global window, canvas, paper, inW, inH, outW, outH, inImage, outImage
    if canvas != None :
        canvas.destroy()
    window.geometry(str(outH) + 'x' + str(outW))
    canvas = Canvas(window, height=outH, width=outW)
    paper = PhotoImage(height=outH, width=outW)
    canvas.create_image((outW / 2, outH / 2), image=paper, state='normal')

    ## 한 픽셀씩 화면에 출력 (열라 느림)
    # for i in range(outH) :
    #     for k in range(outW) :
    #         data = outImage[i][k]
    #         paper.put("#%02x%02x%02x" % (data, data, data), (k, i))
    ## 메모리에 화면(문자열)을 출력해 놓고, 한방에 출력(열라 빠름)
    rgbString = '' # 여기에 전체 픽셀 문자열을 저장할 계획
    step = 1
    for i in range(0, outH, step) :
        tmpString = ''
        for k in range(0, outW, step) :
            data = outImage[i][k]
            tmpString += ' #%02x%02x%02x' % (data, data, data)
        rgbString += '{' + tmpString + '} '
    paper.put(rgbString)


    canvas.pack()

import struct
def saveFile() :
    global window, canvas, paper, inImage, outImage,inW, inH, outW, outH, filename
    saveFp = asksaveasfile(parent=window, mode='wb',
                               defaultextension="*.raw", filetypes=(("RAW파일", "*.raw"), ("모든파일", "*.*")))
    for i in range(outW):
        for k in range(outH):
            saveFp.write( struct.pack('B',outImage[i][k]))

    saveFp.close()

import xlsxwriter
def saveExcelImage() :
    global window, canvas, paper, inImage, outImage, inW, inH, outW, outH, filename
    saveFp = asksaveasfile(parent=window, mode='wb',
                               defaultextension="*.xlsx", filetypes=(("엑셀 파일", "*.xlsx"), ("모든파일", "*.*")))
    xlsxName = saveFp.name

    sheetName = os.path.basename(xlsxName).split(".")[0]
    wb = xlsxwriter.Workbook(xlsxName)
    ws = wb.add_worksheet(sheetName)

    #워크시트의 폭 조절
    ws.set_column(0, outW, 1.0)  # 실제로 약 0.34쯤.
    #워크시트의 높이 조절
    for  r  in range(outH) :
        ws.set_row(r, 9.5)  # 실제로 약 0.35쯤
    # 각 셀마다 색상을 지정하자.
    for rowNum in range(outH) :
        for colNum in range(outW) :
            data = outImage[rowNum][colNum]
            # data 값으로 셀의 배경색을 조절... #000000 ~ #FFFFFF
            if data > 15 :
                hexStr = '#' + hex(data)[2:] * 3
            else :
                hexStr = '#' + ('0' + hex(data)[2:] ) * 3
            ## 셀의 포맷 형식을 준비
            cell_format = wb.add_format()
            cell_format.set_bg_color(hexStr)
            ws.write(rowNum,colNum,'', cell_format)
    wb.close()
    messagebox.showinfo('완료', xlsxName + ' 저장됨')



import pymysql
from tkinter import ttk
IP_ADDR = '192.168.111.141'
DB_NAME = 'machineDB'
TBL_NAME = 'imageTBL'
USER_NAME = 'root'
USER_PASS = '1234'
def loadDB() :
    global window, canvas, paper, inW, inH, outW, outH, inImage, outImage, filename
    global window2, sheet, rows
    con = pymysql.connect(host=IP_ADDR, user=USER_NAME, password=USER_PASS, database=DB_NAME, charset='utf8')
    cur = con.cursor()

    sql = "SELECT id, fName, fType FROM imageTBL"  # ID로 추출하기
    cur.execute(sql)

    rows = cur.fetchall()

    ## 새로운 윈도창 띄우기
    window2 = Toplevel(window)
    sheet = ttk.Treeview(window2, height=10);    sheet.pack()
    descs = cur.description
    colNames = [d[0] for d in descs]
    sheet.column("#0", width=80);
    sheet.heading("#0", text=colNames[0])
    sheet["columns"] = colNames[1:]
    for colName in colNames[1:]:
        sheet.column(colName, width=80);
        sheet.heading(colName, text=colName)
    for row in rows :
        sheet.insert('', 'end', text=row[0], values=row[1:])
    sheet.bind('<Double-1>', sheetDblClick)

    cur.close()
    con.close()

def sheetDblClick(event) :
    global window, canvas, paper, inW, inH, outW, outH, inImage, outImage, filename
    global window2, sheet, rows

    item = sheet.identify('item', event.x, event.y) # 'I001' ....
    entNum = int(item[1:]) - 1  ## 쿼리한 결과 리스트의 순번
    id = rows[entNum][0] ## 선택한 id
    window2.destroy()
    # DB에서 이미지를 다운로드
    con = pymysql.connect(host=IP_ADDR, user=USER_NAME, password=USER_PASS, database=DB_NAME, charset='utf8')
    cur = con.cursor()
    sql = "SELECT fName, image FROM imageTBL WHERE id=" + str(id) # ID로 이미지 추출하기
    cur.execute(sql)
    row = cur.fetchone()
    cur.close()
    con.close()
    import tempfile
    # 임시 폴더
    fname, binData = row
    fullPath = tempfile.gettempdir()+ '/' + fname # 임시경로 + 파일명
    fp = open(fullPath , 'wb') # 폴더를 지정.
    fp.write(binData)
    fp.close()

    if fname.split('.')[1].upper() != 'RAW' :
        messagebox.showinfo('못봄', fname + '은 못봐요.ㅠㅠ')
        return

    filename = fname
    # 파일 --> 메모리
    loadImage(fullPath)
    equalImage()

#################화소영역 처리#################################################################
def embossing() :
    global window, canvas, paper, inW, inH, outW, outH, inImage, outImage
    outImage = []  # 초기화
    # outImage의 크기를 결정
    outH = inH;  outW = inW
    # 빈 메모리 확보 (2차원 리스트)
    for _ in range(outH):
        tmp = []
        for _ in range(outW):
            tmp.append(0)
        outImage.append(tmp)
    #### 영상 처리 알고리즘을 구현 ####
    MSIZE=3
    mask = [ [-1, 0, 0],
             [ 0, 0, 0],
             [ 0, 0, 1],  ]
    # 임시 입력 영상 : inImage보다 2크게
    tmpInImage = []
    for _ in range(inH + 2) :
        tmp = []
        for _ in range(inW + 2) :
            tmp.append(127)
        tmpInImage.append(tmp)
    # 임시 출력 영상 : outImage와 동일
    tmpOutImage = []
    for _ in range(outH ) :
        tmp = []
        for _ in range(outW) :
            tmp.append(0)
        tmpOutImage.append(tmp)
    # 입력 --> 임시입력
    for i in range(inH) :
        for k in range(inW) :
            tmpInImage[i+1][k+1] = inImage[i][k]
    ### 회선 연산 --> 마스크로 쭈욱~~ 긁어서 계산하기...
    for i in range(1, inH) :
        for k in range(1, inW) :
            # 1점을 처리하기. 3x3 반복 처리.  각 위치끼리 곱한후 합계...
            S = 0.0
            for  m  in range(0, MSIZE) :
                for n in range(0, MSIZE) :
                    S += mask[m][n] * tmpInImage[i+(m-1)][k+(n-1)]
            tmpOutImage[i-1][k-1] = S

    ## 마스크의 합계가 0일 경우엔 127 정도를 더한다.
    for i in range(outH):
        for k in range(outW):
            tmpOutImage[i][k] += 127

    ## 임시 출력 --> 출력
    for i in range(outH) :
        for k in range(outW) :
            value = int(tmpOutImage[i][k])
            if value > 255 :
                value = 255
            elif value < 0 :
                value = 0
            outImage[i][k] = value


    ################################
    displayImage()



###################################################################################

def equalImage() :
    global window, canvas, paper, inW, inH, outW, outH, inImage, outImage
    outImage = []  # 초기화
    # outImage의 크기를 결정
    outH = inH;  outW = inW
    # 빈 메모리 확보 (2차원 리스트)
    for _ in range(outH):
        tmp = []
        for _ in range(outW):
            tmp.append(0)
        outImage.append(tmp)
    #### 영상 처리 알고리즘을 구현 ####
    for i in range(inH) :
        for k in range(inW) :
            outImage[i][k] = inImage[i][k]
    ################################
    displayImage()

from tkinter.simpledialog import *
def addImage() :
    global window, canvas, paper, inW, inH, outW, outH, inImage, outImage
    outImage = []  # 초기화
    # outImage의 크기를 결정
    outH = inH;  outW = inW
    # 빈 메모리 확보 (2차원 리스트)
    for _ in range(outH):
        tmp = []
        for _ in range(outW):
            tmp.append(0)
        outImage.append(tmp)
    value = askinteger("밝게할 값", "값 입력")
    #### 영상 처리 알고리즘을 구현 ####
    for i in range(inH) :
        for k in range(inW) :
            if inImage[i][k] + value > 255 :
                outImage[i][k] = 255
            else :
                outImage[i][k] = inImage[i][k] + value
    ################################
    displayImage()

def reverseImage() :
    global window, canvas, paper, inW, inH, outW, outH, inImage, outImage
    outImage = []  # 초기화
    # outImage의 크기를 결정
    outH = inH;  outW = inW
    # 빈 메모리 확보 (2차원 리스트)
    for _ in range(outH):
        tmp = []
        for _ in range(outW):
            tmp.append(0)
        outImage.append(tmp)
    #### 영상 처리 알고리즘을 구현 ####
    for i in range(inH) :
        for k in range(inW) :
            outImage[i][k] = 255 - inImage[i][k]
    ################################
    displayImage()

def mirror1Image() :
    global window, canvas, paper, inW, inH, outW, outH, inImage, outImage
    outImage = []  # 초기화
    # outImage의 크기를 결정
    outH = inH;  outW = inW
    # 빈 메모리 확보 (2차원 리스트)
    for _ in range(outH):
        tmp = []
        for _ in range(outW):
            tmp.append(0)
        outImage.append(tmp)
    #### 영상 처리 알고리즘을 구현 ####
    for i in range(inH) :
        for k in range(inW) :
            outImage[inW-1-i][k] = inImage[i][k]
    ################################
    displayImage()

def bwImage() :
    global window, canvas, paper, inW, inH, outW, outH, inImage, outImage
    outImage = []  # 초기화
    # outImage의 크기를 결정
    outH = inH;  outW = inW
    # 빈 메모리 확보 (2차원 리스트)
    for _ in range(outH):
        tmp = []
        for _ in range(outW):
            tmp.append(0)
        outImage.append(tmp)
    #### 평균값 구하기 ###
    hap = 0
    for i in range(inH) :
        for k in range(inW) :
            hap += inImage[i][k]
    avg = hap // (inH * inW)

    #### 영상 처리 알고리즘을 구현 ####
    for i in range(inH) :
        for k in range(inW) :
            if inImage[i][k] >= avg :
                outImage[i][k] = 255
            else :
                outImage[i][k] = 0
    ################################
    displayImage()

def zoomOut1Image() :
    global window, canvas, paper, inW, inH, outW, outH, inImage, outImage
    outImage = []  # 초기화
    scale = askinteger("축소 값", "값 입력")
    # outImage의 크기를 결정
    outH = inH//scale;  outW = inW//scale
    # 빈 메모리 확보 (2차원 리스트)
    for _ in range(outH):
        tmp = []
        for _ in range(outW):
            tmp.append(0)
        outImage.append(tmp)
    #### 영상 처리 알고리즘을 구현 ####
    for i in range(inH) :
        for k in range(inW) :
            outImage[i//scale][k//scale] = inImage[i][k]
    ################################
    displayImage()

def zoomOut2Image() :
    global window, canvas, paper, inW, inH, outW, outH, inImage, outImage
    outImage = []  # 초기화
    scale = askinteger("축소 값", "값 입력")
    # outImage의 크기를 결정
    outH = inH//scale;  outW = inW//scale
    # 빈 메모리 확보 (2차원 리스트)
    for _ in range(outH):
        tmp = []
        for _ in range(outW):
            tmp.append(0)
        outImage.append(tmp)
    #### 영상 처리 알고리즘을 구현 ####
    for i in range(outH) :
        for k in range(outW) :
            outImage[i][k] = inImage[i*scale][k*scale]
    ################################
    displayImage()

def zoomIn1Image() :
    global window, canvas, paper, inW, inH, outW, outH, inImage, outImage
    outImage = []  # 초기화
    scale = askinteger("축소 값", "값 입력")
    # outImage의 크기를 결정
    outH = inH*scale;  outW = inW*scale
    # 빈 메모리 확보 (2차원 리스트)
    for _ in range(outH):
        tmp = []
        for _ in range(outW):
            tmp.append(0)
        outImage.append(tmp)
    #### 영상 처리 알고리즘을 구현 ####
    for i in range(inH) :
        for k in range(inW) :
            outImage[i*scale][k*scale] = inImage[i][k]
    ################################
    displayImage()


def zoomIn2Image() :
    global window, canvas, paper, inW, inH, outW, outH, inImage, outImage
    outImage = []  # 초기화
    scale = askinteger("축소 값", "값 입력")
    # outImage의 크기를 결정
    outH = inH*scale;  outW = inW*scale
    # 빈 메모리 확보 (2차원 리스트)
    for _ in range(outH):
        tmp = []
        for _ in range(outW):
            tmp.append(0)
        outImage.append(tmp)
    #### 영상 처리 알고리즘을 구현 ####
    for i in range(outH) :
        for k in range(outW) :
            outImage[i][k] = inImage[i//scale][k//scale]
    ################################
    displayImage()

## RAW 영상의 입력 및 출력 평균값 계산
def averageRAW() :
    global window, canvas, paper, inW, inH, outW, outH, inImage, outImage
    rawSum =0
    for i in range(inH) :
        for k in range(inW) :
            rawSum += inImage[i][k]
    inRawAvg = rawSum // (inH * inW) # 입력영상 평균값
    rawSum = 0
    for i in range(outH):
        for k in range(outW):
            rawSum += outImage[i][k]
    outRawAvg = rawSum // (inH * inW)  # 출력영상 평균값

    subWindow = Toplevel(window);    subWindow.geometry('200x100')
    label1 = Label(subWindow, text='입력영상 평균값-->' + str(inRawAvg)); label1.pack()
    label2 = Label(subWindow, text='출력영상 평균값-->' + str(outRawAvg)); label2.pack()
    subWindow.mainloop()

## RAW 영상의 히스토그램 그리기
def histoRAW() :
    global window, canvas, paper, inW, inH, outW, outH, inImage, outImage
    countList = [0] * 256 ; normalList = [0] * 256
    # 빈도수 세기
    for i in range(outH):
        for k in range(outW):
            value = outImage[i][k]
            countList[value] += 1
    # 정규화 시키기 : 정규화된값 = (카운트값 - 최소값) * 최대높이 / (최대값 - 최소값)
    maxValue = max(countList);  minValue = min(countList)
    for i in range(len(countList)) :
        normalList[i] = (countList[i] - minValue) * 256 / (maxValue - minValue)
    # 히스토그램 그리기
    subWindow = Toplevel(window);    subWindow.geometry('256x256')
    subCanvas = Canvas(subWindow, width=256, height=256)
    subPaper = PhotoImage(width=256, height=256)
    subCanvas.create_image( (256/2, 256/2), image=subPaper, state='normal')

    for i in range(0, 256) :
        for k in range(0, int(normalList[i])) :
            if k > 255 :
                break
            data = 0
            subPaper.put('#%02x%02x%02x' % (data, data, data), (i, 255-k))

    subCanvas.pack(expand=1, anchor=CENTER)
    subWindow.mainloop()

import matplotlib.pyplot as plt
def matHistoRAW() :
    global window, canvas, paper, inW, inH, outW, outH, inImage, outImage
    countList = [0] * 256 ;
    # 빈도수 세기
    for i in range(outH):
        for k in range(outW):
            value = outImage[i][k]
            countList[value] += 1
    plt.plot(countList)
    plt.show()

# RAW 이미지(outImage)를 테이블로 행단위로 정보를 저장
'''
DB : machineDB,   Table : colorTBL( id, fname, ftype, x, y, r, g, b)
USE machineDB;  CREATE TABLE colorTBL( id bigint auto_increment PRIMARY KEY, fname VARCHAR(30), ftype VARCHAR(10),
        xSize smallint, ySize smallint, x  smallint, y smallint, r smallint, g smallint, b smallint)
'''
def rawToTable() :
    global window, canvas, paper, inW, inH, outW, outH, inImage, outImage, filename
    global window2, sheet, rows
    if outImage == [] or outImage == None :
        return

    con = pymysql.connect(host=IP_ADDR, user=USER_NAME, password=USER_PASS, database=DB_NAME, charset='utf8')
    cur = con.cursor()

    import random
    if filename == '' or filename == None :
        filename = "image" + random.randint(100000) + ".raw"
    fname = os.path.basename(filename);    ftype = 'RAW'; xSize = outW; ySize = outH


    for i in range(outH) :
        for k in range(outW) :
            sql = "INSERT INTO colorTBL( id, fname, ftype, xSize, ySize, x, y, r, g, b) VALUES (NULL, '" + \
                  fname + "', '" + ftype + "', " + str(xSize) + ", " + str(ySize) + ", "
            r = g = b = outImage[i][k]
            sql += str(i) + ", " + str(k) + ", " + str(r)  + ", " + str(g) + ", " + str(b) + ")"
            cur.execute(sql)

    cur.close()
    con.commit()
    con.close()
    messagebox.showinfo('완료', fname + '이 DB 테이블에 입력 완료')


# 테이블의 행단위로 정보를 RAW 이미지(inImage) 로딩
def tableToRaw() :
    global window, canvas, paper, inW, inH, outW, outH, inImage, outImage, filename
    global window2, sheet, rows
    con = pymysql.connect(host=IP_ADDR, user=USER_NAME, password=USER_PASS, database=DB_NAME, charset='utf8')
    cur = con.cursor()

    sql = "SELECT DISTINCT fName, fType, xSize, ySize FROM colorTBL"  # ID로 추출하기
    cur.execute(sql)

    rows = cur.fetchall()

    ## 새로운 윈도창 띄우기
    window2 = Toplevel(window)
    sheet = ttk.Treeview(window2, height=10);    sheet.pack()
    descs = cur.description
    colNames = [d[0] for d in descs]
    sheet.column("#0", width=80);
    sheet.heading("#0", text=colNames[0])
    sheet["columns"] = colNames[1:]
    for colName in colNames[1:]:
        sheet.column(colName, width=80);
        sheet.heading(colName, text=colName)
    for row in rows :
        sheet.insert('', 'end', text=row[0], values=row[1:])
    sheet.bind('<Double-1>', sheetDblClick2)

    cur.close()
    con.close()

def sheetDblClick2(event) :
    global window, canvas, paper, inW, inH, outW, outH, inImage, outImage, filename
    global window2, sheet, rows

    item = sheet.identify('item', event.x, event.y) # 'I001' ....
    entNum = int(item[1:]) - 1  ## 쿼리한 결과 리스트의 순번
    fileID = rows[entNum][0] ## 선택한 id
    window2.destroy()
    # DB에서 이미지를 다운로드
    con = pymysql.connect(host=IP_ADDR, user=USER_NAME, password=USER_PASS, database=DB_NAME, charset='utf8')
    cur = con.cursor()
    sql = "SELECT x, y, r, g, b FROM colorTBL WHERE fName='" + fileID + "'" # ID로 이미지 추출하기
    cur.execute(sql)

    colorRows = cur.fetchall()
    cur.close()
    con.close()

    ### 입력 영상을 완성 ### (inImage, inW, inH)
    print(rows)
    inW = rows[entNum][2]; inH=rows[entNum][3]
    inImage = []

    # 빈 메모리 확보 (2차원 리스트)
    for _ in range(inH) :
        tmp = []
        for _ in range(inW) :
            tmp.append(0)
        inImage.append(tmp)
    for row in colorRows :
        x, y, r, g, b = row
        inImage[x][y] = r

    # 파일 --> 메모리
    equalImage()


###########################################################
###########CSV 데이터 분석#####################
###########################################################
csvList = []
def openCSVFile() :
    global  csvList
    global window, canvas, paper, inW, inH, outW, outH, inImage, outImage
    filename = askopenfilename(parent=window, filetypes=(("CSV 파일", "*.csv"), ("모든 파일", "*.*")))
    if filename == "" or filename == None:
        return

    loadCSV(filename)
import csv
def loadCSV(fname) :
    global csvList
    with open(fname, 'r', newline='') as filereader:
        csvReader = csv.reader(filereader) # CSV 전용으로 다시 열기
        header_list = next(csvReader)
        csvList.append(header_list)
        for  row_list in csvReader :
            csvList.append(row_list)

    drawSheet(csvList)

sheet = None
def drawSheet(cList) :
    global sheet
    if sheet != None :
        sheet.destroy()

    sheet = ttk.Treeview(window)
    sheet.pack(side=LEFT, fill=Y)

    sheet.column("#0", width=80); sheet.heading("#0", text=cList[0][0])
    sheet["columns"] = cList[0][1:]
    for colName in cList[0][1:]:
        sheet.column(colName, width=80); sheet.heading(colName, text=colName)

    for row in cList[1:] :
        colList = []
        for col in row[1:]:
            colList.append(col)
        sheet.insert('', 'end', text=row[0], values=tuple(colList))

def csvUp10() :
    global csvList
    # cost 열의 위치를 찾자.
    header_list = csvList[0]
    for i in range(len(header_list)) :
        header_list[i] = header_list[i].upper().strip()
    try :
        pos = header_list.index('COST')
    except :
        messagebox.showinfo('메시지', 'COST 열 없음')
        return

    for i in range(1, len(csvList)) :
        row = csvList[i]
        cost = row[pos]
        cost = float(cost[1:])
        cost *= 1.1
        cost_str = "${0:.2f}".format(cost)
        csvList[i][pos] = cost_str

    drawSheet(csvList)

def saveCSVFile() :
    global csvList
    saveFp = asksaveasfile(parent=window, mode='w',
                           defaultextension=".csv", filetypes=(("CSV파일", "*.csv"), ("모든파일", "*.*")))
    with open(saveFp.name, 'w', newline='') as filewriter :
        for row_list in csvList :
            row_str = ','.join(map(str, row_list)) # 리스트 --> ,로 구분된 스트링으로 만들기
            filewriter.writelines(row_str + '\n')

import xlrd
def openExcelFile() :
    global csvList
    global window, canvas, paper, inW, inH, outW, outH, inImage, outImage
    filename = askopenfilename(parent=window, filetypes=(("xls 파일", "*.xls;*.xlsx"), ("모든 파일", "*.*")))
    if filename == "" or filename == None:
        return
    workbook = xlrd.open_workbook(filename)
    sheetCount = workbook.nsheets

    firstYN = True
    for worksheet in workbook.sheets() :
        sRow = worksheet.nrows
        sCol = worksheet.ncols
        for i in range(sRow) :
            if firstYN == True and i == 0 :
                firstYN = False
                pass
            elif firstYN == False and i != 0 :
                pass
            elif firstYN == False and i == 0 :
                print(i, end='   ')
                continue
            tmpList = []
            for k in range(sCol) :
                value = worksheet.cell_value(i,k)
                tmpList.append(value)
            csvList.append(tmpList)

    drawSheet(csvList)

import xlwt
def saveExcelFile() :
    global  csvList
    if csvList == [] or csvList == None :
        return
    saveFp = asksaveasfile(parent=window, mode='w',
                           defaultextension=".xls", filetypes=(("엑셀 파일", "*.xls"), ("모든파일", "*.*")))
    filename = saveFp.name
    workbook = xlwt.Workbook()
    outSheet = workbook.add_sheet('sheet1')
    for i in range(len(csvList)) :
        for k in range(len(csvList[i])) :
            outSheet.write(i,k, csvList[i][k])

    workbook.save(filename)
    messagebox.showinfo('저장완료', filename + '저장됨')

def excelUp10() :
    pass

##############################################
########### 칼라 영상 데이터 처리 ##########
##############################################
from PIL import Image
def loadImageColor(fname) :
    global window, canvas, paper, inW, inH, outW, outH, inImageR, inImageG, inImageB, outImageR, outImageG, outImageB, filename, photo
    inImageR, inImageG, inImageB = [], [], [] # 초기화
    # 파일 크기 계산
    photo = Image.open(fname)
    inW = photo.width;  inH = photo.height
    # 빈 메모리 확보 (2차원 리스트)
    for _ in range(inH) :
        tmp = []
        for _ in range(inW) :
            tmp.append(0)
        inImageR.append(tmp)
    for _ in range(inH) :
        tmp = []
        for _ in range(inW) :
            tmp.append(0)
        inImageG.append(tmp)
    for _ in range(inH) :
        tmp = []
        for _ in range(inW) :
            tmp.append(0)
        inImageB.append(tmp)
    # 파일 --> 메모리로 한개씩 옮기기
    photoRGB = photo.convert('RGB')
    for  i  in  range(inH) :
        for k in range(inW) :
            r, g, b = photoRGB.getpixel((k, i)) #
            inImageR[i][k] = r; inImageG[i][k] = g; inImageB[i][k] = b

    # print(inImageR[100][100],inImageG[100][100],inImageB[100][100])




def  openImageColor() :
    global window, canvas, paper, inW, inH, outW, outH, inImageR, inImageG, inImageB, outImageR, outImageG, outImageB, filename
    filename = askopenfilename(parent=window, filetypes=(("영상 파일", "*.gif;*.jpg;*.png;*.bmp;*.tif"), ("모든 파일", "*.*")))
    if filename == "" or filename == None :
        return
    # 파일 --> 메모리
    loadImageColor(filename)

    # Input --> outPut으로 동일하게 만들기.
    equalImageColor()


def displayImageColor() :
    global window, canvas, paper, inW, inH, outW, outH, inImageR, inImageG, inImageB, outImageR, outImageG, outImageB, filename
    if canvas != None :
        canvas.destroy()
    window.geometry(str(outW) + 'x' + str(outH))
    canvas = Canvas(window, height=outH, width=outW)
    paper = PhotoImage(height=outH, width=outW)
    canvas.create_image((outW / 2, outH / 2), image=paper, state='normal')

    rgbString = '' # 여기에 전체 픽셀 문자열을 저장할 계획
    step = 1
    for i in range(0, outH, step) :
        tmpString = ''
        for k in range(0, outW, step) :
            r, g, b = outImageR[i][k], outImageG[i][k], outImageB[i][k],
            tmpString += ' #%02x%02x%02x' % (r, g, b)
        rgbString += '{' + tmpString + '} '
    paper.put(rgbString)
    canvas.pack()

def equalImageColor() :
    global window, canvas, paper, inW, inH, outW, outH, inImageR, inImageG, inImageB, outImageR, outImageG, outImageB, filename
    outImageR, outImageG, outImageB = [], [], []  # 초기화
    # outImage의 크기를 결정
    outH = inH;  outW = inW
    # 빈 메모리 확보 (2차원 리스트)
    for _ in range(outH):
        tmp = []
        for _ in range(outW):
            tmp.append(0)
        outImageR.append(tmp)
    for _ in range(outH):
        tmp = []
        for _ in range(outW):
            tmp.append(0)
        outImageG.append(tmp)
    for _ in range(outH):
        tmp = []
        for _ in range(outW):
            tmp.append(0)
        outImageB.append(tmp)
    #### 영상 처리 알고리즘을 구현 ####
    for i in range(inH) :
        for k in range(inW) :
            outImageR[i][k] = inImageR[i][k]
            outImageG[i][k] = inImageG[i][k]
            outImageB[i][k] = inImageB[i][k]
    ################################
    displayImageColor()

def addImageColor() :
    global window, canvas, paper, inW, inH, outW, outH, inImageR, inImageG, inImageB, outImageR, outImageG, outImageB, filename
    outImageR, outImageG, outImageB = [], [], []  # 초기화
    # outImage의 크기를 결정
    outH = inH;  outW = inW
    # 빈 메모리 확보 (2차원 리스트)
    for _ in range(outH):
        tmp = []
        for _ in range(outW):
            tmp.append(0)
        outImageR.append(tmp)
    for _ in range(outH):
        tmp = []
        for _ in range(outW):
            tmp.append(0)
        outImageG.append(tmp)
    for _ in range(outH):
        tmp = []
        for _ in range(outW):
            tmp.append(0)
        outImageB.append(tmp)
    value = askinteger("밝게할 값", "값 입력")
    #### 영상 처리 알고리즘을 구현 ####
    for i in range(inH) :
        for k in range(inW) :
            if inImageR[i][k] + value > 255 :
                outImageR[i][k] = 255
            else :
                outImageR[i][k] = inImageR[i][k] + value
            if inImageG[i][k] + value > 255 :
                outImageG[i][k] = 255
            else :
                outImageG[i][k] = inImageG[i][k] + value
            if inImageB[i][k] + value > 255 :
                outImageB[i][k] = 255
            else :
                outImageB[i][k] = inImageB[i][k] + value
    ################################
    displayImageColor()

def reverseImageColor() :
    global window, canvas, paper, inW, inH, outW, outH, inImageR, inImageG, inImageB, outImageR, outImageG, outImageB, filename
    outImageR, outImageG, outImageB = [], [], []  # 초기화
    # outImage의 크기를 결정
    outH = inH;  outW = inW
    # 빈 메모리 확보 (2차원 리스트)
    for _ in range(outH):
        tmp = []
        for _ in range(outW):
            tmp.append(0)
        outImageR.append(tmp)
    for _ in range(outH):
        tmp = []
        for _ in range(outW):
            tmp.append(0)
        outImageG.append(tmp)
    for _ in range(outH):
        tmp = []
        for _ in range(outW):
            tmp.append(0)
        outImageB.append(tmp)
    #### 영상 처리 알고리즘을 구현 ####
    for i in range(inH) :
        for k in range(inW) :
            outImageR[i][k] = 255 - inImageR[i][k]
            outImageG[i][k] = 255 - inImageG[i][k]
            outImageB[i][k] = 255 - inImageB[i][k]
    ################################
    displayImageColor()

def mirror1ImageColor() :
    global window, canvas, paper, inW, inH, outW, outH, inImageR, inImageG, inImageB, outImageR, outImageG, outImageB, filename
    outImageR, outImageG, outImageB = [], [], []  # 초기화
    # outImage의 크기를 결정
    outH = inH;  outW = inW
    # 빈 메모리 확보 (2차원 리스트)
    for _ in range(outH):
        tmp = []
        for _ in range(outW):
            tmp.append(0)
        outImageR.append(tmp)
    for _ in range(outH):
        tmp = []
        for _ in range(outW):
            tmp.append(0)
        outImageG.append(tmp)
    for _ in range(outH):
        tmp = []
        for _ in range(outW):
            tmp.append(0)
        outImageB.append(tmp)
    #### 영상 처리 알고리즘을 구현 ####
    for i in range(inH) :
        for k in range(inW) :
            outImageR[inH-1-i][k] = inImageR[i][k]
            outImageG[inH-1-i][k] = inImageG[i][k]
            outImageB[inH-1-i][k] = inImageB[i][k]
    ################################
    displayImageColor()


def bwImageColor() :
    global window, canvas, paper, inW, inH, outW, outH, inImageR, inImageG, inImageB, outImageR, outImageG, outImageB, filename
    outImageR, outImageG, outImageB = [], [], []  # 초기화
    # outImage의 크기를 결정
    outH = inH;  outW = inW
    # 빈 메모리 확보 (2차원 리스트)
    for _ in range(outH):
        tmp = []
        for _ in range(outW):
            tmp.append(0)
        outImageR.append(tmp)
    for _ in range(outH):
        tmp = []
        for _ in range(outW):
            tmp.append(0)
        outImageG.append(tmp)
    for _ in range(outH):
        tmp = []
        for _ in range(outW):
            tmp.append(0)
        outImageB.append(tmp)
    #### 중위수 추출 ####
    hap = 0
    for i in range(inH) :
        for k in range(inW) :
            tData = (inImageR[i][k] + inImageG[i][k] + inImageB[i][k]) // 3
            hap += tData
    avg = hap // (inW*inH)

    #### 영상 처리 알고리즘을 구현 ####
    for i in range(inH) :
        for k in range(inW) :
            if (inImageR[i][k] + inImageG[i][k] + inImageB[i][k]) // 3 >= avg :
                outImageR[i][k] = outImageG[i][k] = outImageB[i][k] = 255
            else :
                outImageR[i][k] = outImageG[i][k] = outImageB[i][k] = 0
    ################################
    displayImageColor()


def zoomOut2ImageColor() :
    global window, canvas, paper, inW, inH, outW, outH, inImageR, inImageG, inImageB, outImageR, outImageG, outImageB, filename
    outImageR, outImageG, outImageB = [], [], []  # 초기화
    scale = askinteger("축소 값", "값 입력")
    # outImage의 크기를 결정
    outH = inH//scale;  outW = inW//scale
    # 빈 메모리 확보 (2차원 리스트)
    for _ in range(outH):
        tmp = []
        for _ in range(outW):
            tmp.append(0)
        outImageR.append(tmp)
    for _ in range(outH):
        tmp = []
        for _ in range(outW):
            tmp.append(0)
        outImageG.append(tmp)
    for _ in range(outH):
        tmp = []
        for _ in range(outW):
            tmp.append(0)
        outImageB.append(tmp)
    #### 영상 처리 알고리즘을 구현 ####
    for i in range(outH) :
        for k in range(outW) :
            outImageR[i][k] = inImageR[i*scale][k*scale]
            outImageG[i][k] = inImageG[i * scale][k * scale]
            outImageB[i][k] = inImageB[i * scale][k * scale]
    ################################
    displayImageColor()

def zoomIn2ImageColor() :
    global window, canvas, paper, inW, inH, outW, outH, inImageR, inImageG, inImageB, outImageR, outImageG, outImageB, filename
    outImageR, outImageG, outImageB = [], [], []  # 초기화
    scale = askinteger("확대 값", "값 입력")
    # outImage의 크기를 결정
    outH = inH*scale;  outW = inW*scale
    # 빈 메모리 확보 (2차원 리스트)
    for _ in range(outH):
        tmp = []
        for _ in range(outW):
            tmp.append(0)
        outImageR.append(tmp)
    for _ in range(outH):
        tmp = []
        for _ in range(outW):
            tmp.append(0)
        outImageG.append(tmp)
    for _ in range(outH):
        tmp = []
        for _ in range(outW):
            tmp.append(0)
        outImageB.append(tmp)
    #### 영상 처리 알고리즘을 구현 ####
    for i in range(outH) :
        for k in range(outW) :
            outImageR[i][k] = inImageR[i//scale][k//scale]
            outImageG[i][k] = inImageG[i // scale][k // scale]
            outImageB[i][k] = inImageB[i // scale][k // scale]
    ################################
    displayImageColor()

import matplotlib.pyplot as plt
def matHistoColor() :
    global window, canvas, paper, inW, inH, outW, outH, inImage, outImage
    countListR = [0] * 256 ;
    # 빈도수 세기
    for i in range(outH):
        for k in range(outW):
            value = outImageR[i][k]
            countListR[value] += 1
    plt.plot(countListR)

    countListG = [0] * 256;
    for i in range(outH):
        for k in range(outW):
            value = outImageG[i][k]
            countListG[value] += 1
    plt.plot(countListG)

    countListB = [0] * 256;
    for i in range(outH):
        for k in range(outW):
            value = outImageB[i][k]
            countListB[value] += 1
    plt.plot(countListB)
    plt.show()

from PIL import ImageFilter, ImageEnhance
def embossingColor() :
    global window, canvas, paper, inW, inH, outW, outH, inImageR, inImageG, inImageB, outImageR, outImageG, outImageB, filename, photo

    ## Pillow 라이브러리가 제공해주는 메소드(함수)를 사용해서 처리
    photo2 = photo.copy()
    photo2 = photo2.filter(ImageFilter.EMBOSS)

    outImageR, outImageG, outImageB = [], [], []  # 초기화
    # outImage의 크기를 결정
    outH = inH;  outW = inW
    # 빈 메모리 확보 (2차원 리스트)
    for _ in range(outH):
        tmp = []
        for _ in range(outW):
            tmp.append(0)
        outImageR.append(tmp)
    for _ in range(outH):
        tmp = []
        for _ in range(outW):
            tmp.append(0)
        outImageG.append(tmp)
    for _ in range(outH):
        tmp = []
        for _ in range(outW):
            tmp.append(0)
        outImageB.append(tmp)
    #### 영상 처리 알고리즘을 구현 ####
    photoRGB = photo2.convert('RGB')
    for i in range(outH):
        for k in range(outW):
            r, g, b = photoRGB.getpixel((k, i))  #
            outImageR[i][k] = r;
            outImageG[i][k] = g;
            outImageB[i][k] = b
    ################################
    displayImageColor()

##############################################
########### OpenCV 활용 ##########
##############################################
import cv2
def loadImageColorCV2(fname) :
    global window, canvas, paper, inW, inH, outW, outH, inImageR, inImageG, inImageB
    global outImageR, outImageG, outImageB, filename, photo, cvPhoto
    inImageR, inImageG, inImageB = [], [], [] # 초기화

    ##########################################
    ## OpenCV용으로 읽어서 보관 + Pillow용으로 변환
    cvData = cv2.imread(fname)
    cvPhoto = cv2.cvtColor(cvData, cv2.COLOR_BGR2RGB)
    photo = Image.fromarray(cvPhoto)
    #############################################

    # 파일 크기 계산
    inW = photo.width;  inH = photo.height
    # 빈 메모리 확보 (2차원 리스트)
    for _ in range(inH) :
        tmp = []
        for _ in range(inW) :
            tmp.append(0)
        inImageR.append(tmp)
    for _ in range(inH) :
        tmp = []
        for _ in range(inW) :
            tmp.append(0)
        inImageG.append(tmp)
    for _ in range(inH) :
        tmp = []
        for _ in range(inW) :
            tmp.append(0)
        inImageB.append(tmp)
    # 파일 --> 메모리로 한개씩 옮기기
    photoRGB = photo.convert('RGB')
    for  i  in  range(inH) :
        for k in range(inW) :
            r, g, b = photoRGB.getpixel((k, i)) #
            inImageR[i][k] = r; inImageG[i][k] = g; inImageB[i][k] = b

    # print(inImageR[100][100],inImageG[100][100],inImageB[100][100])




def  openOpenCV() :
    global window, canvas, paper, inW, inH, outW, outH, inImageR, inImageG, inImageB
    global outImageR, outImageG, outImageB, filename, photo, cvPhoto

    filename = askopenfilename(parent=window, filetypes=(("영상 파일", "*.jpg;*.png;*.bmp;*.tif"), ("모든 파일", "*.*")))
    if filename == "" or filename == None :
        return
    # 파일 --> 메모리
    loadImageColorCV2(filename)

    # Input --> outPut으로 동일하게 만들기.
    equalImageColor()

import numpy as np
def embossingCV2() :
    global window, canvas, paper, inW, inH, outW, outH, inImageR, inImageG, inImageB
    global outImageR, outImageG, outImageB, filename, photo, cvPhoto

    ####### 엠보싱을 CV2 메소드로 구현하기 --> photo2로 넘기기 ####
    cvPhoto2 = cvPhoto[:] # 복사
    mask = np.zeros((3,3), np.float32);   mask[0][0] = -1 ; mask[2][2] = 1
    cvPhoto2 = cv2.filter2D(cvPhoto2, -1, mask)
    cvPhoto2 += 127
    photo2 = Image.fromarray(cvPhoto2)
    ###########################################################

    ## Pillow 라이브러리가 제공해주는 메소드(함수)를 사용해서 처리
    #photo2 = photo.copy()
    #photo2 = photo2.filter(ImageFilter.EMBOSS)

    toColorImage(photo2)

def grayScaleCV2():
    global window, canvas, paper, inW, inH, outW, outH, inImageR, inImageG, inImageB
    global outImageR, outImageG, outImageB, filename, photo, cvPhoto
    cvPhoto2 = cvPhoto[:]  # 복사
    cvPhoto2 = cv2.cvtColor(cvPhoto2, cv2.COLOR_RGB2GRAY)
    photo2 = Image.fromarray(cvPhoto2)
    ###########################################################

    toColorImage(photo2)

def blurCV2():
    global window, canvas, paper, inW, inH, outW, outH, inImageR, inImageG, inImageB
    global outImageR, outImageG, outImageB, filename, photo, cvPhoto
    cvPhoto2 = cvPhoto[:]  # 복사
    mask = np.ones((9, 9), np.float32)/(9*9);
    cvPhoto2 = cv2.filter2D(cvPhoto2, -1, mask)
    photo2 = Image.fromarray(cvPhoto2)
    ###########################################################

    toColorImage(photo2)

def rotateCV2():
    global window, canvas, paper, inW, inH, outW, outH, inImageR, inImageG, inImageB
    global outImageR, outImageG, outImageB, filename, photo, cvPhoto
    angle = askinteger("각도", "회전각-->")
    cvPhoto2 = cvPhoto[:]  # 복사
    rotate_matrix = cv2.getRotationMatrix2D((inW//2, inH//2), angle, 1) # (기준점, 각도, 스케일)
    cvPhoto2 = cv2.warpAffine(cvPhoto2, rotate_matrix, (inW, inH))
    photo2 = Image.fromarray(cvPhoto2)
    ###########################################################
    cv2.imshow('rotate',cvPhoto2)

# def scaleZICV2():
#     global window, canvas, paper, inW, inH, outW, outH, inImageR, inImageG, inImageB
#     global outImageR, outImageG, outImageB, filename, photo, cvPhoto
#     scale = askfloat("스케일 계수", "확대/축소 값-->", minvalue = 0.1, maxvalue=10.0)
#     cvPhoto2 = cvPhoto[:]  # 복사
#     cvPhoto2 = cv2.resize(cvPhoto2, None, fx=scale, fy=scale)
#     photo2 = Image.fromarray(cvPhoto2)
#     ###########################################################
#     cv2.imshow('scale',cvPhoto2)

def scaleZICV2():
    global window, canvas, paper, inW, inH, outW, outH, inImageR, inImageG, inImageB
    global outImageR, outImageG, outImageB, filename, photo, cvPhoto
    scale = askfloat("스케일 계수", "확대/축소 값-->", minvalue = 1, maxvalue=10)
    cvPhoto2 = cvPhoto[:]  # 복사
    cvPhoto2 = cv2.resize(cvPhoto2, None, fx=scale, fy=scale)
    photo2 = Image.fromarray(cvPhoto2)
    ###########################################################
    toColorImage(cvPhoto2)

def waveVirCV2() :
    global window, canvas, paper, inW, inH, outW, outH, inImageR, inImageG, inImageB
    global outImageR, outImageG, outImageB, filename, photo, cvPhoto
    ####### CV2 메소드로 구현하기 --> photo2로 넘기기 ####
    #cvPhoto2 = cvPhoto[:]  # 복사
    cvPhoto2 = np.zeros(cvPhoto.shape, dtype=cvPhoto.dtype)
    for i in range(inH) :
        for k in range(inW) :

            oy = 0
            if k + ox < inH :
                cvPhoto2[i,k]  = cvPhoto [i, (k+ox)% inW]
            else :
                cvPhoto2[i,k] = 0

    photo2 = Image.fromarray(cvPhoto2)
    ###########################################################
    toColorImage(photo2)


def waveHorCV2() :
    global window, canvas, paper, inW, inH, outW, outH, inImageR, inImageG, inImageB
    global outImageR, outImageG, outImageB, filename, photo, cvPhoto
    ####### CV2 메소드로 구현하기 --> photo2로 넘기기 ####
    #cvPhoto2 = cvPhoto[:]  # 복사
    cvPhoto2 = np.zeros(cvPhoto.shape, dtype=cvPhoto.dtype)
    for i in range(inH) :
        for k in range(inW) :
            oy = int(15.0 * math.sin(2*3.14*k / 180))
            ox = 0
            if i + oy < inH :
                cvPhoto2[i,k]  = cvPhoto [(i+oy)% inH, k]
            else :
                cvPhoto2[i,k] = 0

    photo2 = Image.fromarray(cvPhoto2)
    ###########################################################
    toColorImage(photo2)

def cartoonCV2():
    global window, canvas, paper, inW, inH, outW, outH, inImageR, inImageG, inImageB
    global outImageR, outImageG, outImageB, filename, photo, cvPhoto
    cvPhoto2 = cvPhoto[:]  # 복사
    cvPhoto2 = cv2.cvtColor(cvPhoto2, cv2.COLOR_RGB2GRAY)
    cvPhoto2 = cv2.medianBlur(cvPhoto2, 7)

    edges = cv2.Laplacian(cvPhoto2, cv2.CV_8U, ksize=5)
    ret, mask = cv2.threshold(edges, 100, 255, cv2.THRESH_BINARY_INV)

    cvPhoto2 = cv2.cvtColor(mask, cv2.COLOR_GRAY2RGB)
    photo2 = Image.fromarray(cvPhoto2)
    ###########################################################

    toColorImage(photo2)

def waveCV2() :
    global window, canvas, paper, inW, inH, outW, outH, inImageR, inImageG, inImageB
    global outImageR, outImageG, outImageB, filename, photo, cvPhoto
    ####### CV2 메소드로 구현하기 --> photo2로 넘기기 ####
    #cvPhoto2 = cvPhoto[:]  # 복사
    cvPhoto2 = np.zeros(cvPhoto.shape, dtype=cvPhoto.dtype)
    for i in range(inH) :
        for k in range(inW) :
            ox = int(25.0 * math.sin(2*3.14*i / 180))
            oy = int(15.0 * math.sin(2*3.14*k / 180))
            if i + oy < inH :
                cvPhoto2[i,k] = cvPhoto[(i+oy)%inH,k]
                cvPhoto2[i,k] = cvPhoto[i,(k+ox)%inW]
            else :
                cvPhoto2[i,k] = 0

    photo2 = Image.fromarray(cvPhoto2)
    ###########################################################
    toColorImage(photo2)

def translationCV2():
    global window, canvas, paper, inW, inH, outW, outH, inImageR, inImageG, inImageB
    global outImageR, outImageG, outImageB, filename, photo, cvPhoto
    cvPhoto2 = cvPhoto[:]
    num_rows, num_cols = cvPhoto.shape[:2]
    translation = np.float32([[1, 0, 70], [0, 1, 110]])
    photo_translation = cv2.warpAffine(cvPhoto2, translation, (num_cols, num_rows))
    cv2.imshow('Translation', photo_translation)
    cv2.waitKey()

def scaleZOCV2():
    global window, canvas, paper, inW, inH, outW, outH, inImageR, inImageG, inImageB
    global outImageR, outImageG, outImageB, filename, photo, cvPhoto
    # scale = askfloat("스케일 계수", "확대/축소 값-->", minvalue=0.1)
    # cvPhoto2 = cvPhoto[:]  # 복사
    # cvPhoto2 = cv2.resize(cvPhoto2, None, fx=scale, fy=scale, interpolation=cv2.INTER_AREA)
    # photo2 = Image.fromarray(cvPhoto2)
    # ###########################################################
    # toColorImage(cvPhoto2)

def mirrorCV2():
    global window, canvas, paper, inW, inH, outW, outH, inImageR, inImageG, inImageB
    global outImageR, outImageG, outImageB, filename, photo, cvPhoto
    cvPhoto2 = cvPhoto[::-1]
    outH = inH;  outW = inW
    # 빈 메모리 확보 (2차원 리스트)
    # for _ in range(outH):
    #     tmp = []
    #     for _ in range(outW):
    #         tmp.append(0)
    #     cvPhoto2.append(tmp)
    # for i in range(inH) :
    #     for k in range(inW) :
    #         cvPhoto2[inH-1-i][k] = cvPhoto[i][k]
    photo2 = Image.fromarray(cvPhoto2)
    ###########################################################
    toColorImage(photo2)

def blur2CV2():
    global window, canvas, paper, inW, inH, outW, outH, inImageR, inImageG, inImageB
    global outImageR, outImageG, outImageB, filename, photo, cvPhoto
    cvPhoto2 = cvPhoto[:]  # 복사
    value = askinteger("블러링 계수", "블러링 값-->")
    mask = np.ones((value, value), np.float32)/(value*value);
    cvPhoto2 = cv2.filter2D(cvPhoto2, -1, mask)
    photo2 = Image.fromarray(cvPhoto2)
    ###########################################################

    toColorImage(photo2)

def toColorImage(photo2):
    global window, canvas, paper, inW, inH, outW, outH, inImageR, inImageG, inImageB
    global outImageR, outImageG, outImageB, filename, photo, cvPhoto
    outImageR, outImageG, outImageB = [], [], []  # 초기화
    # outImage의 크기를 결정
    outH = inH;
    outW = inW
    # 빈 메모리 확보 (2차원 리스트)
    for _ in range(outH):
        tmp = []
        for _ in range(outW):
            tmp.append(0)
        outImageR.append(tmp)
    for _ in range(outH):
        tmp = []
        for _ in range(outW):
            tmp.append(0)
        outImageG.append(tmp)
    for _ in range(outH):
        tmp = []
        for _ in range(outW):
            tmp.append(0)
        outImageB.append(tmp)
    #### 영상 처리 알고리즘을 구현 ####
    photoRGB = photo2.convert('RGB')
    for i in range(outH):
        for k in range(outW):
            r, g, b = photoRGB.getpixel((k, i))  #
            outImageR[i][k] = r;
            outImageG[i][k] = g;
            outImageB[i][k] = b
    ################################
    displayImageColor()

## 전역 변수 선언부
window, canvas, paper = [None] * 3;  inW, inH, outW, outH = [200] * 4
inImage, outImage = [], []
filename = None
inImageR, inImageG, inImageB, outImageR, outImageG, outImageB = [],[],[],[],[],[],
photo, cvPhoto = None, None # Pillow용, OpenCV용

## 메인 코드부

window = Tk(); window.title('빅데이터 분석/처리 통합 툴 (Ver 0.06)')
window.geometry('800x500')

## 메뉴 추가하기

mainMenu = Menu(window) # 메뉴 전체 껍질
window.config(menu=mainMenu)

rawDataMenu = Menu(mainMenu)
mainMenu.add_cascade(label="RAW 데이터 분석", menu=rawDataMenu)

rawfileMenu = Menu(rawDataMenu)
rawDataMenu.add_cascade(label="파일", menu=rawfileMenu)
rawfileMenu.add_command(label="파일에서 열기", command=openImage)
rawfileMenu.add_command(label="DB에서 불러오기", command=loadDB)
rawfileMenu.add_separator()
rawfileMenu.add_command(label="저장", command=saveFile)
rawfileMenu.add_command(label="엑셀로 저장", command=saveExcelImage)

rawImage1Menu = Menu(rawDataMenu)
rawDataMenu.add_cascade(label="영상처리1", menu=rawImage1Menu)
rawImage1Menu.add_command(label="밝게하기", command=addImage)
rawImage1Menu.add_command(label="반전하기", command=reverseImage)
rawImage1Menu.add_command(label="미러링(상하)", command=mirror1Image)
rawImage1Menu.add_command(label="흑백", command=bwImage)
rawImage1Menu.add_command(label="축소(포워딩)", command=zoomOut1Image)
rawImage1Menu.add_command(label="축소(백워딩)", command=zoomOut2Image)
rawImage1Menu.add_command(label="확대(포워딩)", command=zoomIn1Image)
rawImage1Menu.add_command(label="확대(백워딩)", command=zoomIn2Image)

rawImage2Menu = Menu(rawDataMenu)
rawDataMenu.add_cascade(label="영상처리2", menu=rawImage2Menu)
rawImage2Menu.add_command(label="엠보싱", command=embossing)

rawStatMenu = Menu(rawDataMenu)
rawDataMenu.add_cascade(label="통계 분석", menu=rawStatMenu)
rawStatMenu.add_command(label="평균값", command=averageRAW)
rawStatMenu.add_command(label="히스토그램", command=histoRAW)
rawStatMenu.add_command(label="히스토그램(MatPlotLib)", command=matHistoRAW)

rawDBMenu = Menu(rawDataMenu)
rawDataMenu.add_cascade(label="테이블 입출력", menu=rawDBMenu)
rawDBMenu.add_command(label="RAW->테이블", command=rawToTable)
rawDBMenu.add_command(label="테이블->RAW", command=tableToRaw)
#################################
ecDataMenu = Menu(mainMenu)
mainMenu.add_cascade(label="엑셀(CSV) 데이터 분석", menu=ecDataMenu)

csvfileMenu = Menu(ecDataMenu)
ecDataMenu.add_cascade(label="CSV 분석", menu=csvfileMenu)
csvfileMenu.add_command(label="CSV 파일 열기", command=openCSVFile)
csvfileMenu.add_command(label="CSV 파일 저장", command=saveCSVFile)
csvfileMenu.add_separator()
csvfileMenu.add_command(label="가격10%인상", command=csvUp10)

excelfileMenu = Menu(ecDataMenu)
ecDataMenu.add_cascade(label="엑셀 분석", menu=excelfileMenu)
excelfileMenu.add_command(label="엑셀 파일 열기", command=openExcelFile)
excelfileMenu.add_command(label="엑셀 파일 저장", command=saveExcelFile)
excelfileMenu.add_separator()
excelfileMenu.add_command(label="가격10%인상", command=excelUp10)

#################################
colorDataMenu = Menu(mainMenu)
mainMenu.add_cascade(label="칼라 영상 데이터 분석", menu=colorDataMenu)
colorDataMenu.add_command(label="칼라 영상 파일 열기", command=openImageColor)

colorDataMenu.add_command(label="밝게하기", command=addImageColor)
colorDataMenu.add_command(label="반전하기", command=reverseImageColor)
colorDataMenu.add_command(label="미러링(상하)", command=mirror1ImageColor)
colorDataMenu.add_command(label="흑백", command=bwImageColor)
colorDataMenu.add_command(label="축소(백워딩)", command=zoomOut2ImageColor)
colorDataMenu.add_command(label="확대(백워딩)", command=zoomIn2ImageColor)
colorDataMenu.add_command(label="히스토그램(MatPlotLib)", command=matHistoColor)
colorDataMenu.add_separator()
colorDataMenu.add_command(label="엠보싱", command=embossingColor)

#################################
openCVMenu = Menu(mainMenu)
mainMenu.add_cascade(label="OpenCV(머신러닝)", menu=openCVMenu)
openCVMenu.add_command(label="이미지 열기(OpenCV)", command=openOpenCV)
openCVMenu.add_separator()
openCVMenu.add_command(label="엠보싱(OpenCV)", command=embossingCV2)
openCVMenu.add_command(label="Greyscale(OpenCV)", command=grayScaleCV2)
openCVMenu.add_command(label="블러링(OpenCV)", command=blurCV2)
openCVMenu.add_separator()
openCVMenu.add_command(label="회전(OpenCV)", command=rotateCV2)
openCVMenu.add_command(label="확대/축소(OpenCV)", command=scaleZICV2)
openCVMenu.add_separator()
openCVMenu.add_command(label="수직 웨이브(OpenCV)", command=waveVirCV2)
openCVMenu.add_command(label="수평 웨이브(OpenCV)", command=waveHorCV2)
openCVMenu.add_command(label="카툰화(OpenCV)", command=cartoonCV2)
openCVMenu.add_separator()
openCVMenu.add_command(label="웨이브(OpenCV)", command=waveCV2)
openCVMenu.add_command(label="이동(OpenCV)", command=translationCV2)
openCVMenu.add_command(label="축소(OpenCV)", command=scaleZOCV2)
openCVMenu.add_command(label="미러링(OpenCV)", command=mirrorCV2)
openCVMenu.add_command(label="블러링 입력(OpenCV)", command=blur2CV2)

window.mainloop()