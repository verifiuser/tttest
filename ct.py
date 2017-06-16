# -*- coding: utf-8 -*-

import xlrd
import random
import sys
import hashlib
import daterange
import csvparser
import switch
import underid


reload(sys)
sys.setdefaultencoding('utf-8')


class Ct:

    def __init__(self):

        self.ct_id = ""
        self.lbc_office_id = ""
        self.ct_gid = ""
        self.ct_uni_id3 = ""
        self.ct_category = 0
        self.ct_target_type = 0
        self.ct_action_id = 0
        self.ct_action_type = 0
        self.ct_product_ids = ""
        self.ct_action_date = ""
        self.ct_update_at = ""
        self.ct_del_flag = 0

        dr = daterange.DataRange()
        self.sw = switch.Switch()
        self.udi = underid.UnderId()
        self.cs = csvparser.CsvParser()

        book = xlrd.open_workbook('landscape_dummy_data_definition_file.xls')
        sheets = book.sheets()
        self.s1 = sheets[0]

        self.header = 1
        self.category_list = [1, 2, 3, 4, 5]
        self.target_list = [0, 1]
        self.action_list = [3008, 3152]
        self.product_list = ["US001", "DO002", "DISH001", "DM001", "DM001/DS001", "DM002", "DO001", "ETC", "LBC001",
                             "LBC001/DS001", "LBC001/ETC", "LBC001/TM001", "TM001", "TM001/DM001/DS001", "US002"]
        self.type_list = [0, 1, 2, 3, 4, 5, 6, 7]
        self.ad_list = dr.random_date(span_list=(dr.date_span(start_year=2010, end_year=2017)))
        self.ud_list = dr.random_date_time(span_list=(dr.date_span(start_year=2001, end_year=2009)))

        self.rows = []

    @staticmethod
    def main():

        for row in xrange(vct.header, vct.s1.nrows):

            vct.lbc_office_id = str(vct.s1.cell(row, 0).value)
            vct.ct_gid = str(vct.s1.cell(row, 2).value)
            gn_count = int(vct.s1.cell(vct.sw.case(vct.s1.cell(row, 3).value), 5).value)

            for i in xrange(gn_count):

                vct.ct_id = vct.ct_gid + vct.udi.calculation(count=i)
                vct.ct_uni_id3 = hashlib.md5(vct.ct_id + vct.lbc_office_id).hexdigest()
                vct.ct_category = random.choice(vct.category_list)
                vct.ct_target_type = random.choice(vct.target_list)
                vct.ct_action_id = random.choice(vct.action_list)
                vct.ct_action_type = random.choice(vct.type_list)
                vct.ct_product_ids = random.choice(vct.product_list)
                vct.ct_action_date = random.choice(vct.ad_list)
                vct.ct_update_at = random.choice(vct.ud_list)

                vct.rows.append(
                    [
                        vct.ct_id, vct.lbc_office_id, vct.ct_gid, vct.ct_uni_id3, vct.ct_category, vct.ct_target_type,
                        vct.ct_action_id, vct.ct_action_type, vct.ct_product_ids, vct.ct_action_date, vct.ct_update_at,
                        vct.ct_del_flag
                    ]
                )
        vct.cs.savedata(rows=vct.rows, name='ct', extension='.csv', encode='utf-8')

if __name__ == "__main__":

    vct = Ct()
    vct.main()
    del vct
