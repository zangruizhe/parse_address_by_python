#!/usr/bin/python
# -*- coding: UTF-8 -*-

import traceback
import os, sys
# sys.path.append("python_package/xlrd/lib/python")
# sys.path.append("python_package/xlwt/lib")
sys.path.append(os.path.join('python_package', 'xlrd', 'lib', 'python'))
sys.path.append(os.path.join('python_package', 'xlwt', 'lib'))
sys.path.append(os.path.join('python_package'))

import xlrd
import xlwt
from xlrd import open_workbook

import ntpath
import json

# create logger
#----------------------------------------------------------------------
import logging

log = logging.getLogger('python_logger')
log.setLevel(logging.DEBUG)

fh = logging.FileHandler('out.log', 'w')
fh.setLevel(logging.DEBUG)
# create console handler with a higher log level
ch = logging.StreamHandler()
ch.setLevel(logging.INFO)
# create formatter and add it to the handlers
# formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
# 2015-08-28 17:01:57,662 - simple_example - ERROR - error message

# formatter = logging.Formatter('%(asctime)s %(levelname)-8s %(filename)s:%(lineno)-4d: %(message)s')
formatter = logging.Formatter('%(asctime)s %(levelname)-2s %(lineno)-4d: %(message)s')

fh.setFormatter(formatter)
ch.setFormatter(formatter)
# add the handlers to the logger
log.addHandler(fh)
log.addHandler(ch)
#----------------------------------------------------------------------


#----------------------------------------------------------------------
def open_file(path):
    """
    Open and read an Excel file
    """
    book = xlrd.open_workbook(path)
    # log.info(number of sheets
    log.info(book.nsheets)
    # log.info(sheet names
    log.info(book.sheet_names())
    # get the first worksheet
    first_sheet = book.sheet_by_index(0)
    # read a row
    log.info(first_sheet.row_values(0))
    # read a cell
    cell = first_sheet.cell(0,0)
    log.info(cell)
    log.info(cell.value)
    # read a row slice
    log.info(first_sheet.row_slice(rowx=0,
                                start_colx=0,
                                end_colx=2))
#----------------------------------------------------------------------



class Arm(object):
    def __init__(self, dest_name, addr, phone_num, src_name):
        self.dest_name = dest_name
        self.addr = addr
        self.phone_num = phone_num
        self.src_name = src_name

    def __str__(self):
        return("Arm object:\n"
               "  dest_name = {0}\n"
               "  addr = {1}\n"
               "  phone_num = {2}\n"
               "  src_name = {3}\n"
               .format((self.dest_name).encode('GBK'), (self.addr).encode('GBK'), unicode(self.phone_num)[:-2], (self.src_name).encode('GBK')))



def ParseXls(path):
  log.info("ParseXls:%s", path)
  wb = open_workbook(path)
  for sheet in wb.sheets():
    number_of_rows = sheet.nrows
    number_of_columns = sheet.ncols
    if (number_of_rows <= 0 or number_of_columns <= 0):
      continue

    items = []

    for row in range(number_of_rows):
      values = []
      values.append(sheet.cell(row,0).value.replace(u'\xa0', "").replace(u" ", "").replace(u'\uff0c',"").replace(u",",""))
      values.append(sheet.cell(row,1).value.replace(u'\xa0', "").replace(u" ", "").replace(u'\uff0c',"").replace(u",",""))
      values.append(sheet.cell(row,2).value)
      values.append(sheet.cell(row,3).value.replace(u'\xa0', "").replace(u" ", "").replace(u'\uff0c',"").replace(u",",""))
      item = Arm(*values)
      items.append(item)

  log.info("get account_info number:{}".format(len(items)))

  try:
    for item in items:
      log.info(item)
      break
  except Exception:
    log.error("Got exception on ParseXls:%s", traceback.format_exc() )
  return items

def WriteXls(path, element_list, element_info_list):
  workbook = xlwt.Workbook()
  sheet = workbook.add_sheet("right address")

  head_list = [u"订单号",u"商品名称",u"单位",u"数量",u"单价",u"重量",u"SKU编码", u"SKU名称", u"客户名称*", u"备注",u"收件人姓名",u"收件人省",u"收件人市",u"收件人区",u"收件人地址",u"收件人邮编",u"收件人电话",u"收件人手机",u"收件人邮箱",u"发件人姓名",u"发件人省",u"发件人市",u"发件人区",u"发件人地址",u"发件人邮编",u"发件人电话",u"发件人手机",u"发件人邮箱",u"扩展单号",u"批次号",u"大头笔",u"面单号",u"代收货款",u"到付款",u"网点ID"]
  for index, item in enumerate(head_list) :
    sheet.write(0, index, item)

  for row in range(1, len(element_list) + 1):
    sheet.write(row, head_list.index(u"数量"), 1)
    sheet.write(row, head_list.index(u"代收货款"), element_info_list[1])
    sheet.write(row, head_list.index(u"客户名称*"), element_info_list[2].decode('GBK'))
    sheet.write(row, head_list.index(u"备注"), element_info_list[0] + element_info_list[3] + "%04d"%row)
    sheet.write(row, head_list.index(u"收件人姓名"), element_list[row - 1 ].dest_name)
    sheet.write(row, head_list.index(u"收件人地址"), element_list[row - 1].addr)
    sheet.write(row, head_list.index(u"收件人电话"), unicode(element_list[row - 1].phone_num)[:-2])
    sheet.write(row, head_list.index(u"发件人姓名"), element_list[row - 1].src_name)
  workbook.save(path)
  log.info("finish WriteXls:{}, row numbers:{}".format(path, len(element_list)))


def GetFileList(path):
  from os import listdir
  from os.path import isfile, join
  onlyfiles = [ join(path, f) for f in listdir(path) if isfile(join(path, f))]
  return onlyfiles

def PathLeaf(path):
  head, tail = ntpath.split(path)
  return tail or ntpath.basename(head)

def BuildProvinceInfoList(path = ""):
  log.info("BuildProvinceInfoList path:" + path)
  file_list = GetFileList(path)

  province_info_list = []

  for file_name in file_list:
    f = open(file_name, 'r')
    for line in f:
      jeson_decode_list = json.loads(line)
      province_info_list.extend(jeson_decode_list)
    f.close()
  return province_info_list


def RebuildAddrByDict(src_addr, province_dict_list, debug = False):
  dst_addr = ""
  find_province = False
  find_city = False

  src_addr = src_addr.replace(u'\xa0', "").replace(u" ", "").replace(u'\uff0c',"").replace(u",","")

  province_idx = src_addr.find(u"省")
  city_idx = src_addr.find(u"市")
  if (province_idx != -1 and city_idx != -1) :
    return src_addr

  for province_dict in province_dict_list:

    if (src_addr[:2] != province_dict["province_name"][:2]) :
      continue

    province_name_idx_addr = src_addr.find(province_dict["province_name"][:-1])
    province_name_idx_dict = src_addr.find(province_dict["province_name"][:2])
    if (province_name_idx_addr != -1 or province_name_idx_dict != -1) :
      # find the province
      # if (debug) :
      #   log.info("find the province id:", province_dict["province_id"]

      find_province = True

      if (len(province_dict["city_name"]) < 2 or province_dict["city_name"] == u"自治区直辖县级行政区划") :
        continue

      province_name_idx = -1
      city_name_idx_addr = -1
      city_name_idx_dict = -1
      # find the province
      if (province_name_idx_addr != -1) :
        province_name_idx = province_name_idx_addr + len(province_dict["province_name"][:-1])
      elif (province_name_idx_dict != -1) :
        province_name_idx = province_name_idx_dict + len(province_dict["province_name"][:2])

      city_name_idx_addr = src_addr[province_name_idx : province_name_idx + len(province_dict["city_name"])].find(province_dict["city_name"][:-1])
      city_name_idx_dict = src_addr[province_name_idx : province_name_idx + 4].find(province_dict["city_name"][:2])

      # if (debug) : log.info(src_addr[province_name_idx :]
      if (city_name_idx_addr != -1 or city_name_idx_dict != -1) :

        city_name_idx_addr = src_addr.find(province_dict["city_name"][:-1])
        city_name_idx_dict = src_addr.find(province_dict["city_name"][:2])
        if (debug) :
          log.info("city_name_idx_addr:", city_name_idx_addr, "city_name_idx_dict:", city_name_idx_dict)
          log.info("src", src_addr)
          log.info("find the province_name:", province_dict["province_name"])
          log.info("find the city_name:", province_dict["city_name"])
          log.info("find the county_name:", province_dict["county_name"])
          log.info("find the city id:", province_dict["city_id"])

        dst_addr = province_dict["province_name"]

        if (u"市" in province_dict["city_name"]) :
          dst_addr = dst_addr + province_dict["city_name"]
        else:
          dst_addr = dst_addr + province_dict["city_name"] + u"/市"

        if (city_name_idx_addr != -1) :
          if (src_addr[city_name_idx_addr + len(province_dict["city_name"][:-1])] == province_dict["city_name"][-1]) :
            dst_addr = dst_addr + src_addr[city_name_idx_addr + len(province_dict["city_name"]) : ]
          else :
            dst_addr = dst_addr + src_addr[city_name_idx_addr + len(province_dict["city_name"]) - 1 : ]
          return dst_addr
        elif (city_name_idx_dict != -1) :
          if (src_addr[city_name_idx_dict + 2] == province_dict["city_name"][-1]) :
            dst_addr = dst_addr + src_addr[city_name_idx_dict + 3 : ]
          else :
            dst_addr = dst_addr + src_addr[city_name_idx_dict + 2 : ]
          return dst_addr
      else :
        continue

    else :
      continue

  if (find_province) :
    if (debug) : log.info("find province but can not find city so find county name")
    for province_dict in province_dict_list:

      province_name_idx_addr = src_addr.find(province_dict["province_name"][:-1])
      province_name_idx_dict = src_addr.find(province_dict["province_name"][:2])
      if (province_name_idx_addr != -1 or province_name_idx_dict != -1) :
        # find the province
        province_name_idx = -1
        county_name_idx_addr = -1
        county_name_idx_dict = -1
        # find the province
        if (province_name_idx_addr != -1) :
          province_name_idx = province_name_idx_addr + len(province_dict["province_name"][:-1])
        elif (province_name_idx_dict != -1) :
          province_name_idx = province_name_idx_dict + len(province_dict["province_name"][:2])

        # county_name_idx_addr = src_addr[province_name_idx : ].find(province_dict["county_name"][:-1])
        # county_name_idx_dict = src_addr[province_name_idx : ].find(province_dict["county_name"][:2])
        county_name_idx_addr = src_addr[province_name_idx : province_name_idx + len(province_dict["county_name"])].find(province_dict["county_name"][:-1])
        county_name_idx_dict = src_addr[province_name_idx : province_name_idx + 3 ].find(province_dict["county_name"][:2])

        if (county_name_idx_addr != -1 or county_name_idx_dict != -1) :
          county_name_idx_addr = src_addr.find(province_dict["county_name"][:-1])
          county_name_idx_dict = src_addr.find(province_dict["county_name"][:2])
          if (debug) :
            log.info("find the country name")
            log.info("src", src_addr)
            log.info("find the province_name:", province_dict["province_name"])
            log.info("find the city_name:", province_dict["city_name"])
            log.info("find the county_name:", province_dict["county_name"])
            log.info("find the city id:", province_dict["city_id"])

          dst_addr = province_dict["province_name"]

          if (u"市" in province_dict["city_name"]) :
            dst_addr = dst_addr + province_dict["city_name"]
          else:
            dst_addr = dst_addr + province_dict["city_name"] + u"/市"

          if (county_name_idx_addr != -1) :

            dst_addr = dst_addr + src_addr[county_name_idx_addr : ]

          elif (county_name_idx_dict != -1):

            if (src_addr[county_name_idx_dict + 2] == province_dict["county_name"][-1]) :
              dst_addr = dst_addr + province_dict["county_name"] + src_addr[county_name_idx_dict + 3 : ]
            else :
              dst_addr = dst_addr + province_dict["county_name"] + src_addr[county_name_idx_dict + 2 : ]

          return dst_addr
        else :
          continue
      else :
        continue
  else :
    pass

  return src_addr



def Start() :
  province_info_list = BuildProvinceInfoList(os.path.join('positionJson', 'town'))
  log.info("province_info_list len:%d", len(province_info_list))

  log.info(province_info_list[0]['town_name'][:-1].encode('GBK'))

  doc_list = GetFileList("document")
  for doc in doc_list:
    account_info_list = ParseXls(doc)
    log.info("get account_info_list:{}".format(len(account_info_list)))

    illegal_account_list = []
    for item in account_info_list:
      if (item.addr.find(u"省") != -1 and item.addr.find(u"市") != -1):
        continue
      else:
        item.addr = RebuildAddrByDict(item.addr, province_info_list)
        illegal_account_list.append(item)

    log.info("illegal_account_list len:%d", len(illegal_account_list))

    account_list_can_not_prase = []

    for item in illegal_account_list:
      if (item.addr.find(u"省") != -1 and item.addr.find(u"市") != -1) :
        # if (item.addr.find(u"新疆") != -1) :
        #   log.info(item.addr.encode('UTF-8')
        # log.info("account_list_prase",item.addr.encode('UTF-8')
        pass
      else:
        # log.info("account_list_can_not_prase",item.addr
        account_list_can_not_prase.append(item)

    log.info("account_list_can_not_prase len:%d", len(account_list_can_not_prase))

    doc_name_info_list = PathLeaf(doc).split("-")

    WriteXls(os.path.join('result_document', 'result_' + PathLeaf(doc)), account_info_list, doc_name_info_list)
    #end

def CheckTheDict(province_name, city_name, province_dict_list):
  for province_dict in province_dict_list:
    if (province_dict["province_name"] == province_name):
      if (province_dict["city_name"] == city_name) :
        log.info("find the province info:", province_dict)
        return
      else:
        continue
    else:
      continue
  log.info("can not find the province info")
  return




if __name__ == "__main__":
  # log.info("province_info_list len:", len(province_info_list)
  # CheckTheDict(u"北京市", u"崇文区",  province_info_list)
  # 
  debug = False
  if (debug):
    province_info_list = BuildProvinceInfoList("positionJson/town")
    test_addr = u"新疆维吾尔自治区阿苏克拜城县拜城镇交通路46号(艺苑小区)13栋1单元402室"
    log.info(RebuildAddrByDict(test_addr, province_info_list, True))
    test_addr = u"内蒙古省阿拉善盟阿拉善右旗阿右旗地税局"
    log.info(RebuildAddrByDict(test_addr, province_info_list, True))
    test_addr = u"内蒙古省兴安盟科尔沁右翼中旗西哲里木镇前索根嘎查"
    log.info(RebuildAddrByDict(test_addr, province_info_list, True))
    test_addr = u"河南省清丰县马村乡孟卜村"
    log.info(RebuildAddrByDict(test_addr, province_info_list, True))
    test_addr = u"，新疆 塔城地区 沙湾县大泉乡中泉村3巷10号"
    log.info(RebuildAddrByDict(test_addr, province_info_list, True))
    test_addr = u"新疆维吾尔族阿克苏库车县五一中路南建转盘"
    log.info(RebuildAddrByDict(test_addr, province_info_list, True))
  else :
    try:
      Start()
    except Exception:
      log.error("Got exception on ParseXls:%s", traceback.format_exc() )

  raw_input("press Enter to exit")
  """
  jeson_string = '[{"province_id":110,"province_name":"北京市","city_id":"110100000000","city_name":"市辖区","county_id":"110101000000","county_name":"东城区","town_id":"110101001000","town_name":"东华门街道办事处"},{"province_id":110,"province_name":"北京市","city_id":"110100000000","city_name":"市辖区","county_id":"110101000000","county_name":"东城区","town_id":"110101002000","town_name":"景山街道办事处"},{"province_id":110,"province_name":"北京市","city_id":"110100000000","city_name":"市辖区","county_id":"110101000000","county_name":"东城区","town_id":"110101001000","town_name":"东华门街道办事处"},{"province_id":110,"province_name":"北京市","city_id":"110100000000","city_name":"市辖区","county_id":"110101000000","county_name":"东城区","town_id":"110101002000","town_name":"景山街道办事处"}]'
  jeson_decode = json.loads(jeson_string)
  log.info(jeson_decode, type(jeson_decode)
  for jeson_content in jeson_decode:
    log.info(jeson_content['town_name'].encode('UTF-8')
    """

