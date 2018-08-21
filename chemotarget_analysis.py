import win32com
from win32com.client import Dispatch, constants, gencache
import pandas as pd
import os
from shutil import copyfile
import time
from pandas import ExcelWriter
import numpy as np
from collections import defaultdict, OrderedDict

def add_basic_informmation(doc, informdict, barcode):
    nowtime = time.strftime('%Y/%m/%d', time.localtime())
    partnername = pd.read_excel('E:\\化疗套餐报告自动化\\肿瘤个体化化疗套餐项目报告自动化资料\\capitalname.xlsx', index_col=0)
    partnerlist = partnername.index.tolist()
    doctable = doc.Tables[0]
    doctable.Cell(1, 2).Range.Text = str(barcode)
    if informdict[barcode]['采集时间'] != '':
        doctable.Cell(1, 4).Range.Text = informdict[barcode]['采集时间'].strftime('%Y/%m/%d')
    else:
        doctable.Cell(1, 4).Range.Text = ''
    doctable.Cell(2, 2).Range.Text = informdict[barcode]['样本号']
    doctable.Cell(2, 4).Range.Text = informdict[barcode]['录入时间'].strftime('%Y/%m/%d')
    doctable.Cell(3, 4).Range.Text = nowtime
    # doctable.Cell(3,4).Range.Text = '2018/04/11' #报告时间
    doctable.Cell(7, 3).Range.Text = informdict[barcode]['姓名']
    doctable.Cell(9, 2).Range.Text = informdict[barcode]['性别']
    doctable.Cell(9, 4).Range.Text = informdict[barcode]['临床诊断']
    doctable.Cell(10, 2).Range.Text = str(informdict[barcode]['岁']) + '岁'

    for partner in partnerlist:
        if partner in informdict[barcode]['医院名称']:
            doctable.Cell(10, 4).Range.Text = ''
        else:
            doctable.Cell(10, 4).Range.Text = informdict[barcode]['医院名称']

    doctable.Cell(11, 2).Range.Text = informdict[barcode]['身份证号']
    doctable.Cell(11, 4).Range.Text = informdict[barcode]['送检医生']
    doctable.Cell(12, 2).Range.Text = informdict[barcode]['标本类型']
    doctable.Cell(12, 4).Range.Text = informdict[barcode]['病人号']
    doctable.Cell(13, 2).Range.Text = '正常'  #标本状态
    doctable.Cell(13, 4).Range.Text = informdict[barcode]['病理编号']
    # doc.Tables[3].Cell(1, 2).Range.Text = informdict[barcode]['检测者']
    doc.Tables[3].Cell(1, 4).Range.Text = informdict[barcode]['审核时间'].strftime('%Y/%m/%d')
    # doc.Tables[3].Cell(2, 2).Range.Text = '杨慧敏'
    # doc.Tables[3].Cell(2, 2).Range.Text = informdict[barcode]['审核者']
    doc.Tables[3].Cell(2, 4).Range.Text = nowtime
    return doc

def add_experiment_result(doc, persondata, wapp):
    drugorder = sort_by_drug(persondata['关联药物'].tolist())
    data_grouped = persondata.groupby(by='关联药物')
    rownum = 2
    for drugname in drugorder:
        evrygroup = data_grouped.get_group(drugname)
        minnum = len(evrygroup)
        metaanalysis = meta_analysis(psdata=evrygroup, drugname=drugname)
        if minnum >1:
            doc.Tables[1].Cell(rownum, 1).Select()      #合并第一列
            wapp.Selection.MoveDown(Unit=5, Count=minnum-1, Extend=1)
            wapp.Selection.Cells.Merge()
            doc.Tables[1].Cell(rownum, 1).Range.Text = drugname

            # doc.Tables[1].Cell(rownum, 2).Select()   #合并第一列
            # wapp.Selection.MoveDown(Unit=5, Count=minnum - 1, Extend=1)
            # wapp.Selection.Cells.Merge()
            # doc.Tables[1].Cell(rownum, 2).Range.Text = evrygroup['肿瘤'].tolist()[0]

            doc.Tables[1].Cell(rownum, 5).Select()  #合并最后一列，写入综合分析结果
            wapp.Selection.MoveDown(Unit=5, Count=minnum - 1, Extend=1)
            wapp.Selection.Cells.Merge()
            doc.Tables[1].Cell(rownum,5).Range.Text = metaanalysis[0]

        elif minnum == 1:
            doc.Tables[1].Cell(rownum, 1).Range.Text = drugname
            # doc.Tables[1].Cell(rownum, 2).Range.Text = evrygroup['肿瘤'].tolist()[0]
            doc.Tables[1].Cell(rownum, 5).Range.Text = metaanalysis[0]

        for minproject in evrygroup.index.tolist():
            doc.Tables[1].Cell(rownum, 2).Range.Text = persondata.loc[minproject, '检测项目']
            doc.Tables[1].Cell(rownum, 3).Range.Text = persondata.loc[minproject, '检测结果']
            doc.Tables[1].Cell(rownum, 4).Range.Text = persondata.loc[minproject, '意义']

            if rownum <= len(persondata)+1:
                rownum +=1
    return doc

def meta_analysis(psdata, drugname):
    resultanalysis = []
    drugtypelist = [i for i in set(psdata['药物类型'].tolist())]
    if len(drugtypelist) == 1:
        for drugtypename, drugtypegroup in psdata.groupby(by='药物类型'):
            if len(set(drugtypegroup['意义'].tolist())) > 1:
                if drugtypename == '药物治疗':
                    mindescription = '该检测个体对%s药物治疗敏感性降低，建议综合考虑毒副作用适当调整剂量使用。' % drugname.replace('/', '、')
                elif drugtypename == '毒副作用':
                    mindescription = '该检测个体常规剂量下%s药物治疗毒副作用风险相对增加。' % drugname.replace('/', '、')

            elif len(set(drugtypegroup['意义'].tolist())) == 1:
                if drugtypename == '药物治疗' or drugtypename == '药物治疗和毒副作用':
                    if '敏感性降低' in drugtypegroup['意义'].tolist()[0]:
                        mindescription = '该检测个体对%s相对不敏感。' % drugname.replace('/', '、')
                    else:
                        mindescription = '该检测个体对%s%s。' % (drugname.replace('/', '、'), drugtypegroup['意义'].tolist()[0])
                elif drugtypename == '毒副作用':
                    mindescription = '该检测个体常规剂量下%s药物治疗%s。' % (drugname.replace('/', '、'), drugtypegroup['意义'].tolist()[0])

            resultanalysis.append(mindescription)

    elif len(drugtypelist) > 1:
        mindict = {}
        for drugtypename, drugtypegroup in psdata.groupby(by='药物类型'):
            if drugtypename == '药物治疗' or drugtypename == '药物治疗和毒副作用':
                if len(set(drugtypegroup['意义'].tolist())) ==1:
                    if len(drugtypegroup) > 1 and '敏感性降低' in drugtypegroup['意义'].tolist()[0]:
                        mindict['药物治疗'] = '该检测个体对%s相对不敏感，' % drugname.replace('/', '、')
                    else:
                        mindict['药物治疗'] = '该检测个体对%s%s，' % (drugname.replace('/', '、'), drugtypegroup['意义'].tolist()[0])
                elif len(set(drugtypegroup['意义'].tolist())) >1:
                    mindict['药物治疗'] = '该检测个体对%s药物治疗敏感性降低，' % drugname.replace('/', '、')
                    mindict['补充'] = '建议综合考虑毒副作用适当调整剂量使用。'

            elif drugtypename == '毒副作用':
                if len(set(drugtypegroup['意义'].tolist())) == 1:
                    mindict['毒副作用'] = '常规剂量下%s' % drugtypegroup['意义'].tolist()[0]
                elif len(set(drugtypegroup['意义'].tolist())) >1:
                    mindict['毒副作用'] = '常规剂量下毒副作用风险相对增加'

        if len(mindict.keys()) == 3:
            newdes = mindict['药物治疗'] + mindict['毒副作用'] + '，' + mindict['补充']
        elif len(mindict.keys()) == 2 and '毒副作用' in mindict.keys():
            newdes = mindict['药物治疗'] + mindict['毒副作用'] + '。'
        elif len(minldict.keys()) == 2 and '毒副作用' not in mindict.keys():
            newdes = mindict['药物治疗'] + mindict['补充']

        resultanalysis.append(newdes)
    # resultanalysis = [i for i in set(resultanalysis)]
    return resultanalysis

def sort_by_drug(analysislist):
    # sortdict = defaultdict(list)
    sortdict = {}
    drugsortlist = pd.read_excel('E:\\化疗套餐报告自动化\\肿瘤个体化化疗套餐项目报告自动化资料\\药物顺序表.xlsx', index_col=0, constants={'顺序号':str})
    drugsortdict = drugsortlist.to_dict(orient='index')
    for item in analysislist:
        for i in drugsortdict.keys():
            if i in item:
                sortdict[item] = drugsortdict[i]['顺序号']
    sort_result = sorted(sortdict.items(), key=lambda sortdict:sortdict[1])
    sortedanalysis = [i[0] for i in sort_result]
    return sortedanalysis


def main(Expresultfiles):
    reporttemplate = u'E:\\化疗套餐报告自动化\\肿瘤个体化化疗套餐项目报告自动化资料\\化疗套餐模板_合并综述_20180803.docx'  # 读取报告模板
    reporttemplate_2 = u'E:\\化疗套餐报告自动化\\肿瘤个体化化疗套餐项目报告自动化资料\\个体化报告模板_B5模板_合并单元格.docx'  # 读取报告模板
    backgroudfile = pd.read_excel('E:\\化疗套餐报告自动化\\肿瘤个体化化疗套餐项目报告自动化资料\\单项化疗数据库_V5_20180725.xlsx')  # 背景资料文件
    for file in Expresultfiles:
        Expresultfile = pd.ExcelFile(file)  # 读取需要出具报告的受试者信息表

        sampleinform = Expresultfile.parse(sheetname='基本信息', index_col='条码', converters={'身份证号': str})
        sampleinform.fillna('', inplace=True)

        informdict = sampleinform.to_dict(orient='index')  # 将信息表转化成dict，以条形码为key
        for eachsample in informdict.keys():  # 打印正在生成的检测者
            reportname = os.path.join(os.getcwd(), '%s_%s_%s.docx'%(eachsample, informdict[eachsample]['姓名'], informdict[eachsample]['医院名称']))
            reportname2 = os.path.join(os.getcwd(), '%s_%s_%s.pdf'%(eachsample, informdict[eachsample]['姓名'], informdict[eachsample]['医院名称']))
            print('正在生成【%s_%s】的报告，请稍等！检测项目是：%s'%(eachsample, informdict[eachsample]['姓名'], informdict[eachsample]['检验目的名称']))
            copyfile(reporttemplate, reportname)

            data_person = Expresultfile.parse(sheetname=str(eachsample))  # 读取每个人的检测结果

            w = win32com.client.Dispatch('Word.Application')
            w.Visible = 0
            w.DisplayAlerts = 0
            doc = w.Documents.Open(FileName=reportname)
            doc = add_basic_informmation(doc=doc, informdict=informdict, barcode=eachsample)  # 添加每个受试者的个人信息

            personlist = []
            for minproject in data_person.index.tolist():
                pro_inform = backgroudfile[backgroudfile['检测结果'] == data_person.loc[minproject, '审核人结果']][
                    backgroudfile['检测项目'] == data_person.loc[minproject, '项目名称']]
                for mer in pro_inform.index.tolist():
                    if data_person['癌种'].tolist()[0].strip() in pro_inform.loc[mer, :]['肿瘤']:
                        personlist.append(pro_inform.loc[mer, :])
                    else:
                        pass
                # personlist.append(pro_inform)
            personinform = pd.concat(personlist, axis=1).T    #每个检测者的检测结果

            for rownum in range(0,len(personinform)-1, 1):    #根据结果的行数增加表格中的行数
                doc.Tables[1].Rows.Add()

            for per_pro in data_person['项目名称'].tolist():
                if per_pro not in personinform['检测项目'].tolist():
                    print('该受试者的检测项目有缺失：%s' % per_pro)
                    break
            else:
                doc = add_experiment_result(doc=doc, persondata=personinform, wapp=w)
                if np.isnan(data_person['HE染色结果'][0]) == False:
                    doc.Tables[2].Cell(1, 1).Range.Text = '注: HE染色结果分析其肿瘤组织含量约为%s。' % (
                        format(data_person['HE染色结果'][0], '.0%'))

                backgenelist = personinform['背景资料'].tolist()
                subtabnum = 0
                for tabnum in range(5, doc.Tables.Count):
                    if doc.Tables[tabnum - subtabnum].Rows[2].Range.Text.split('\r')[0] not in backgenelist:
                        doc.Tables[tabnum - subtabnum].Delete()
                        subtabnum += 1
                    else:
                        pass
                doc.SaveAs(reportname2, 17)
                doc.Close()

                # 如果是平邑县医院，则另生成B5版报告
                if '平邑' in informdict[eachsample]['医院名称']:
                    reportname_B5 = os.path.join(os.getcwd(), '%s_%s_%s_B5.docx' % (
                    eachsample, informdict[eachsample]['姓名'], informdict[eachsample]['医院名称']))
                    reportname_B5_2 = os.path.join(os.getcwd(), '%s_%s_%s_B5.pdf' % (
                    eachsample, informdict[eachsample]['姓名'], informdict[eachsample]['医院名称']))
                    copyfile(reporttemplate_2, reportname_B5)
                    w = win32com.client.Dispatch('Word.Application')
                    w.Visible = 0
                    w.DisplayAlerts = 0
                    doc2 = w.Documents.Open(FileName=reportname_B5)
                    doc2 = add_basic_informmation(doc=doc, informdict=informdict, barcode=eachsample)  # 添加每个受试者的个人信息
                    if np.isnan(data_person['HE染色结果'][0]) == False:
                        doc2.Tables[2].Cell(1, 1).Range.Text = '注: HE染色结果分析其肿瘤组织含量约为%s。' % (
                            format(data_person['HE染色结果'][0], '.0%'))
                    rownumber = 1
                    for everyitem in com_analysis_result:
                        doc2.Tables[3].Rows[rownumber].Range.Text = str(rownumber) + '）' + everyitem
                        if rownumber < len(com_analysis_result):
                            rownumber += 1
                            doc2.Tables[3].Rows.Add()
                    doc2.SaveAs(reportname_B5_2, 17)
                    doc2.Close()


if __name__ == '__main__':
    # Expresultfiles = ['个体化20180608-1_合并测试 - 副本.xlsm']
    # main(Expresultfiles=Expresultfiles)
    main()