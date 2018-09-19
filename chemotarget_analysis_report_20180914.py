import win32com
from win32com.client import Dispatch, constants, gencache
import pandas as pd
import os
from shutil import copyfile
import time
from pandas import ExcelWriter
import numpy as np
from collections import defaultdict, OrderedDict
from glob import glob
import math

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
    doctable.Cell(7, 3).Range.Text = informdict[barcode]['姓名']
    doctable.Cell(9, 2).Range.Text = informdict[barcode]['性别']
    doctable.Cell(9, 4).Range.Text = informdict[barcode]['临床诊断']
    doctable.Cell(10, 2).Range.Text = str(informdict[barcode]['岁']) + '岁'

    for partner in partnerlist:     #部分代理商不显示在医院名称中
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
    doc.Tables[3].Cell(1, 4).Range.Text = informdict[barcode]['审核时间'].strftime('%Y/%m/%d')
    doc.Tables[3].Cell(2, 4).Range.Text = nowtime
    return doc

def extract_result(backgroudfile, data_person, target_cancer, chemo_cancer):
    personlist = []
    for minproject in data_person.index.tolist():
        pro_inform = backgroudfile[(backgroudfile['检测项目'] == data_person.loc[minproject, '项目名称']) &
                                   (backgroudfile['检测结果'] == data_person.loc[minproject, '审核人结果'])]         # 将每个检测样本的检测项目和结果对应的数据库中的信息提取出来,但是未区分肿瘤

        for typename, group in pd.groupby(pro_inform, by='检测项目类型'):
            if typename == '靶向':
                tar_re = group[group['癌种'].str.contains('/'+target_cancer)]
                personlist.append(tar_re)
            elif typename == '化疗':
                chem_re = group[group['癌种'].str.contains('/'+chemo_cancer)]
                personlist.append(chem_re)
    personinform = pd.concat(personlist)  # 每个检测者的检测结果
    return personinform

def analysis_personresult(persondata, target_cancer, chemo_cancer):   #对检测者的检测结果按照药物进行分析
    drugorder = sort_by_drug(persondata['关联药物'].tolist())
    data_grouped = persondata.groupby(by='关联药物')
    metadict = OrderedDict()
    tarlist = []
    for drugname in drugorder:
        evrygroup = data_grouped.get_group(drugname)    #每一种药对应的所有项目
        if '靶向' in evrygroup['检测项目类型'].tolist():
            merdata = meta_analysis_targetdrug_new(targetdata=evrygroup, drugname=drugname)
            tarlist.append(merdata)
        elif '化疗' in evrygroup['检测项目类型'].tolist():
            merdict = meta_analysis_chemo(chemodata=evrygroup, drugname=drugname)
            metadict.update(merdict)
    #对神经胶质瘤特殊处理
    if len(tarlist) != 0:
        tardict = drugmerge_analysis(tarlist)
        if target_cancer == '神经胶质瘤' and 'TERT基因突变分析' in tardict['/']['druggroup']['检测项目'].tolist():
            dmgroup = tardict['/']['druggroup']
            if '预后较好' not in tardict['替莫唑胺']['druggroup']['意义'].__str__():
                dmgroup['意义'][dmgroup['检测项目'] == 'TERT基因突变分析'] = '结合IDH检测结果分析预后欠佳,突变常见于原发性胶质母细胞瘤和少突星形细胞瘤'
                dmgroup['药物对应意义'][dmgroup['检测项目'] == 'TERT基因突变分析'] = '预后欠佳'
                dmgroup['意义'][dmgroup['检测项目'] == 'ATRX蛋白表达水平分析'] = '结合IDH检测结果分析预后欠佳'
                dmgroup['药物对应意义'][dmgroup['检测项目'] == 'ATRX蛋白表达水平分析'] = '预后欠佳'
                tardict['替莫唑胺']['meta_con'] = tardict['替莫唑胺']['meta_con'].replace('替莫唑胺', '替莫唑胺预后较差，')
            elif '预后较好' in tardict['替莫唑胺']['druggroup']['意义'].__str__():
                dmgroup['意义'][dmgroup['检测项目'] == 'TERT基因突变分析'] = '结合IDH检测结果分析预后较好,突变常见于原发性胶质母细胞瘤和少突星形细胞瘤'
                dmgroup['药物对应意义'][dmgroup['检测项目'] == 'TERT基因突变分析'] = '预后较好'
                dmgroup['意义'][dmgroup['检测项目'] == 'ATRX蛋白表达水平分析'] = '结合IDH检测结果分析预后较好'
                dmgroup['药物对应意义'][dmgroup['检测项目'] == 'ATRX蛋白表达水平分析'] = '预后较好'
                tardict['替莫唑胺']['meta_con'] = tardict['替莫唑胺']['meta_con'].replace('替莫唑胺', '替莫唑胺预后较好，')
            if '预后较好' in tardict['/']['druggroup']['药物对应意义'].__str__():
                tardict['/']['meta_con'] = '预后较好'
            else:
                tardict['/']['meta_con'] = '预后较差'

        #处理同时存在于化疗和靶向的药物
        elif target_cancer == '结直肠癌' and '微卫星不稳定性(MSI)分析' in tardict['氟尿嘧啶类/卡培他滨']['druggroup']['检测项目'].tolist() and '氟尿嘧啶类/卡培他滨' in metadict.keys():
            fngr = pd.concat([metadict['氟尿嘧啶类/卡培他滨']['druggroup'], tardict['氟尿嘧啶类/卡培他滨']['druggroup']])
            tar_yy = [i for i in set(tardict['氟尿嘧啶类/卡培他滨']['druggroup']['药物对应意义'].tolist())]
            che_yy = metadict['氟尿嘧啶类/卡培他滨']['druggroup']['药物对应意义'].tolist()
            tar_msi = [i for i in set(tardict['氟尿嘧啶类/卡培他滨']['druggroup']['意义'].tolist())]
            if tar_yy[0] not in che_yy:
                meta_tc = metadict['氟尿嘧啶类/卡培他滨']['meta_con'].replace('该检测个体常规剂量下', '该检测个体%s,'%tar_msi[0])
                if '药物治疗敏感性降低，' not in meta_tc:
                    meta_tc = meta_tc.replace('药物治疗相对敏感,', '药物治疗敏感性降低，')
            else:
                meta_tc = metadict['氟尿嘧啶类/卡培他滨']['meta_con'].replace('该检测个体常规剂量下', '该检测个体%s，'%tar_msi[0])
            metadict['氟尿嘧啶类/卡培他滨']={'druggroup':fngr, 'meta_con':meta_tc}
            tardict.pop('氟尿嘧啶类/卡培他滨')

        metadict.update(tardict)
    return metadict

def drugmerge_analysis(drug_sl):
    tar_rs_merge = pd.concat(drug_sl)
    mergedrug_dict = OrderedDict()
    for (protype, minmeta), group in tar_rs_merge.groupby(by=['对应项目合并', '综合分析结果']):
        dl = [i for i in set(group['关联药物'].tolist())]   #需要合并的药物
        dl_sort = sort_by_drug(dl)
        mergedrug = '/'.join(dl_sort)
        group['合并后药物'] = mergedrug
        if mergedrug not in mergedrug_dict.keys():
            setgroup = group.drop_duplicates(['检测项目', '检测结果','药物对应意义'])     #去除重复项
            setgroup = setgroup.sort('检测项目')
            meta_con = '该检测个体对%s%s' %(mergedrug, minmeta)
            mergedrug_dict[mergedrug] = {'druggroup':setgroup, 'minnum':len(setgroup), 'meta_con':meta_con}
        else:
            pass
    return mergedrug_dict

def meta_analysis_chemo(chemodata, drugname):
    drugmetadict = OrderedDict()
    drugtypelist = [i for i in set(chemodata['药物类型'].tolist())]
    if len(drugtypelist) == 1:
        for drugtypename, drugtypegroup in chemodata.groupby(by='药物类型'):
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

            drugmetadict[drugname] = {'druggroup':chemodata, 'minnum':len(chemodata), 'meta_con':mindescription}

    elif len(drugtypelist) > 1:
        mindict = {}
        for drugtypename, drugtypegroup in chemodata.groupby(by='药物类型'):
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

        drugmetadict[drugname] = {'druggroup':chemodata,'minnum':len(chemodata),'meta_con':newdes}
    return drugmetadict

def meta_analysis_targetdrug_new(targetdata, drugname):
    mindict = {}
    senlist = [i for i in set(targetdata['药物对应意义'].tolist())]
    deslist = [i for i in set(targetdata['意义'].tolist())]
    if len(senlist) == 1 and len(deslist) == 1:
        mindiscription = '%s'%deslist[0]
    elif len(senlist) == 1 and len(deslist)>1:
        mindiscription = '%s'%senlist[0]
    elif len(senlist) > 1:
        if '任一敏感则敏感' in targetdata['判断规则'].tolist().__str__():
            mindiscription = '药物治疗相对敏感'
        elif '全敏感则敏感' in targetdata['判断规则'].tolist().__str__():
            mindiscription = '药物治疗相对不敏感'
        elif 'Kras为主' in targetdata['判断规则'].tolist().__str__():
            if 'K-ras' in targetdata['检测项目'].tolist().__str__():
                kras_yiyi = targetdata[targetdata['检测项目'] == 'K-ras code12,13突变分析']['药物对应意义']
                mindiscription = '%s'% kras_yiyi
            else:
                mindiscription = '药物治疗相对敏感'
    # mindict[drugname] = {'meta_con':mindiscription}
    pro_drug = targetdata['检测项目'].tolist()
    pro_drug.sort()
    xiangmu = '/'.join(pro_drug)
    targetdata['综合分析结果'] = mindiscription
    targetdata['对应项目合并'] = xiangmu
    return targetdata


def add_metaresult(alldict, doc, wapp, resdata):
    rownum = 2
    # resdata_grouped = resdata.groupby('关联药物')
    for drug in alldict.keys():
        regroup = alldict[drug]['druggroup']
        minnum = alldict[drug]['minnum']
        if minnum > 1:
            doc.Tables[1].Cell(rownum, 1).Select()  # 合并第一列，写入样品名称
            wapp.Selection.MoveDown(Unit=5, Count=minnum - 1, Extend=1)
            wapp.Selection.Cells.Merge()
            doc.Tables[1].Cell(rownum, 1).Range.Text = drug

            doc.Tables[1].Cell(rownum, 5).Select()  # 合并最后一列，写入综合分析结果
            wapp.Selection.MoveDown(Unit=5, Count=minnum - 1, Extend=1)
            wapp.Selection.Cells.Merge()
            doc.Tables[1].Cell(rownum, 5).Range.Text = alldict[drug]['meta_con'].replace('/', '、')

        else:
            doc.Tables[1].Cell(rownum, 1).Range.Text = drug
            doc.Tables[1].Cell(rownum, 5).Range.Text = alldict[drug]['meta_con'].replace('/', '、')

        for minproject in regroup.index.tolist():
            doc.Tables[1].Cell(rownum, 2).Range.Text = regroup.loc[minproject, '检测项目']
            doc.Tables[1].Cell(rownum, 3).Range.Text = regroup.loc[minproject, '检测结果']
            doc.Tables[1].Cell(rownum, 4).Range.Text = regroup.loc[minproject, '意义']

            if rownum <= len(resdata) + 1:
                rownum += 1
    return doc

def sort_by_drug(analysislist):
    sortdict = {}
    drugsortlist = pd.read_excel('E:\\化疗靶向库文件\\药物顺序表.xlsx', index_col=0, constants={'顺序号':str})
    drugsortdict = drugsortlist.to_dict(orient='index')
    for item in analysislist:
        for i in drugsortdict.keys():
            if i in item:
                sortdict[item] = drugsortdict[i]['顺序号']
    sort_result = sorted(sortdict.items(), key=lambda sortdict:sortdict[1])
    sortedanalysis = [i[0] for i in sort_result]
    return sortedanalysis

def add_picture(doc, persencode):
    imlist = [i for i in glob('%s\\*.jpg'%persencode)]
    doc.Tables[4].Cell(1,1).Range.Text = str(persencode) + ' 检测结果附图/Pictures：'
    for dnum in range(0, math.ceil(len(imlist)/2)*2-2, 1):
        doc.Tables[4].Rows.Add()

    imtable_row = 2
    col_num = 1
    for img in imlist:
        img_name = img.split('\\')[-1].replace('.jpg', '')
        doc.Tables[4].Cell(imtable_row,col_num).Range.Text = img_name
        shape = doc.Tables[4].Cell(imtable_row+1, col_num).Range.InlineShapes.AddPicture(img, LinkToFile=False, SaveWithDocument=True)
        shape.Height, shape.Width = 140, 210
        if col_num == 1:
            col_num = 3
        else:
            col_num = 1
            imtable_row +=2

def main(Expresultfiles):
    os.chdir('E:\\化疗靶向库文件\\测试结果')
    backgroudfile = pd.read_excel('E:\\化疗靶向库文件\\化疗靶向用药数据库_靶向药物单列.xlsx')  # 背景资料文件

    for file in Expresultfiles:
        Expresultfile = pd.ExcelFile(file)  # 读取需要出具报告的受试者信息表
        sampleinform = Expresultfile.parse(sheetname='基本信息', index_col='条码', converters={'身份证号': str})
        sampleinform.fillna('', inplace=True)
        informdict = sampleinform.to_dict(orient='index')  # 将信息表转化成dict，以条形码为key

        for eachsample in informdict.keys():  # 打印正在生成的检测者
            print('正在生成【%s_%s】的报告，请稍等！检测项目是：%s' % (eachsample, informdict[eachsample]['姓名'], informdict[eachsample]['检验目的名称']))
            data_person = Expresultfile.parse(sheetname=str(eachsample))  # 解析检测者的检测结果
            target_cancer = data_person['靶向癌种'].tolist()[0].strip()  # 靶向项目对应的癌种
            chemo_cancer = data_person['靶向癌种'].tolist()[0].strip()  # 化疗项目对应的癌种
            personinform = extract_result(backgroudfile, data_person, target_cancer=target_cancer, chemo_cancer=chemo_cancer)  # 提取出检测者的检测结果对应的背景信息
            metadict_all = analysis_personresult(persondata=personinform, target_cancer=target_cancer, chemo_cancer=chemo_cancer)  # 分析检测者的检测结果,生成药物-综合分析-对应检测项目数量的字典

            if len(data_person['蛋白表达强度'].dropna()) != 0:    #对于需要展示蛋白强度的项目，用蛋白强度的数值替换掉原来的阳性
                protein_pers = {}
                for i in data_person['蛋白表达强度'].dropna().index:
                    protein_pers[data_person.loc[i, '项目名称']] = data_person.loc[i, '蛋白表达强度']
                for drugkey in metadict_all.keys():
                    for proteinpro in protein_pers.keys():
                        if proteinpro in metadict_all[drugkey]['druggroup']['检测项目'].tolist():
                            metadict_all[drugkey]['druggroup']['检测结果'] = protein_pers[proteinpro]

            if '是' in personinform['是否插入图片'].tolist():
                reporttemplate = u'E:\\化疗靶向库文件\\个体化报告模板_V1_插入图片版.docx'  # 读取报告模板
                reporttemplate_2 = u'E:\\化疗靶向库文件\\个体化报告模板_V1_插入图片版_B5版本.docx'  # 读取报告模板
            else:
                reporttemplate = u'E:\\化疗靶向库文件\\个体化报告模板_V1.docx'  # 读取报告模板
                reporttemplate_2 = u'E:\\化疗靶向库文件\\个体化报告模板_V1_B5版本.docx'  # 读取报告模板

            reportname = os.path.join(os.getcwd(), '%s_%s_%s.docx'%(eachsample, informdict[eachsample]['姓名'], informdict[eachsample]['医院名称']))
            reportname2 = os.path.join(os.getcwd(), '%s_%s_%s.pdf'%(eachsample, informdict[eachsample]['姓名'], informdict[eachsample]['医院名称']))
            copyfile(reporttemplate, reportname)

            w = win32com.client.Dispatch('Word.Application')
            w.Visible = 0
            w.DisplayAlerts = 0
            doc = w.Documents.Open(FileName=reportname)
            doc = add_basic_informmation(doc=doc, informdict=informdict, barcode=eachsample)  # 添加每个受试者的个人信息
            rownum_all = 0
            for drug in metadict_all.keys():
                rownum_all += metadict_all[drug]['minnum']
            for rownum in range(0,rownum_all-1, 1):    #根据结果的行数增加表格中的行数
                doc.Tables[1].Rows.Add()
            doc = add_metaresult(alldict=metadict_all, wapp=w, doc=doc, resdata=personinform)     #将检测结果写入到word中
            if np.isnan(data_person['HE染色结果'][0]) == False:     #如果有肿瘤组织含量结果则写入报告中，如果没有则不写
                doc.Tables[2].Cell(1, 1).Range.Text = '注: HE染色结果分析其肿瘤组织含量约为%s。' % (
                    format(data_person['HE染色结果'][0], '.0%'))

            doc = add_picture(doc=doc, persencode=eachsample)   #添加图片

            backgenelist = personinform['背景资料'].tolist()    #背景资料基因
            subtabnum = 0
            for tabnum in range(4, doc.Tables.Count):
                if doc.Tables[tabnum - subtabnum].Rows[2].Range.Text.split('\r')[0] not in backgenelist:    #将不在基因列表中的背景资料删除，存在的留下
                    doc.Tables[tabnum - subtabnum].Delete()
                    subtabnum += 1
                else:
                    pass

            misspro = [i for i in data_person['项目名称'].tolist() if i not in personinform['检测项目'].tolist()]
            if len(misspro) ==0:
                doc.SaveAs(reportname2, 17)
            else:
                print('该受试者的检测项目缺失--%s项：%s' %(len(misspro), misspro))
            doc.Close()

            # 如果是平邑县医院，则另生成B5版报告
            if '平邑' in informdict[eachsample]['医院名称']:
                reportname_B5 = os.path.join(os.getcwd(), '%s_%s_%s_B5.docx' % (eachsample, informdict[eachsample]['姓名'], informdict[eachsample]['医院名称']))
                reportname_B5_2 = os.path.join(os.getcwd(), '%s_%s_%s_B5.pdf' % (eachsample, informdict[eachsample]['姓名'], informdict[eachsample]['医院名称']))
                copyfile(reporttemplate_2, reportname_B5)
                w = win32com.client.Dispatch('Word.Application')
                w.Visible = 0
                w.DisplayAlerts = 0
                doc_b5 = w.Documents.Open(FileName=reportname_B5)
                doc_b5 = add_basic_informmation(doc=doc_b5, informdict=informdict, barcode=eachsample)  # 添加每个受试者的个人信息
                for rownum in range(0, rownum_all-1, 1):  # 根据结果的行数增加表格中的行数
                    doc_b5.Tables[1].Rows.Add()

                doc_b5 = add_metaresult(alldict=metadict_all, wapp=w, doc=doc_b5, resdata=personinform)  # 将检测结果写入到word中
                if np.isnan(data_person['HE染色结果'][0]) == False:  # 如果有肿瘤组织含量结果则写入报告中，如果没有则不写
                    doc_b5.Tables[2].Cell(1, 1).Range.Text = '注: HE染色结果分析其肿瘤组织含量约为%s。' % (format(data_person['HE染色结果'][0], '.0%'))

                for bgtab in range(5, doc_b5.Tables.Count):
                    doc_b5.Tables[bgtab].Delete()
                if len(misspro) == 0:
                    doc_b5.SaveAs(reportname_B5_2, 17)
                else:
                    print('该受试者的检测项目缺失--%s项：%s' % (len(misspro), misspro))
                doc_b5.Close()

if __name__ == '__main__':
    # Expresultfiles = ['化疗靶向测试-test.xlsm']
    # main(Expresultfiles=Expresultfiles)
    main()