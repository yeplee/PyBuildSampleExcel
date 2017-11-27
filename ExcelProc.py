#!usr/bin/env python
# Version: 10.Nov.2017
# Author: ying.fu@gemalto.com
# Filename: ExcelProc.py

# --------------------------------------------------------------------
# The ExcelProc.py tool is
#
# --------------------------------------------------------------------

# coding = 'UTF-8'
import os, sys, getopt, copy
import xlrd, xlwt
import ordereddict

default_out_fname = 'excelProc.log'
default_pcbvender = 'Multek'

default_remark_normal = ' not mounted to eval board'
default_remark_eval = 'Using normal board come from item 1'

MORNAL_ALIGN_QUAN = 16
EVAL_ALIGN_QUAN = 4

class QuantitySample():
    
    def __init__(self, cur_build_name = '', cur_variant = '', cur_PN = ''):        
        #self.qty_dict = {}
        self.qty_dict = ordereddict.OrderedDict()
        self.cur_build_name = cur_build_name
        self.cur_variant = cur_variant
        self.cur_PN = cur_PN

    # append quantity with item name(skey)
    def setQty(self, skey, qty):
        self.qty_dict.setdefault(skey, qty)

    # retrive quantity by item name(skey)
    def getQty(self,skey):
        return self.qty_dict.get(skey)

    def setBuildName(self, build_name):
        self.cur_build_name = build_name

    def getBuildName(self):
        return self.cur_build_name
    
    def setVariant(self, variant):
        self.cur_variant = variant
    
    def getVariant(self):
        return self.cur_variant
    
    def setPN(self, PN):
        self.cur_PN = PN
    
    def getPN(self):
        return self.cur_PN

    def shrinkQtyDict(self):
        for each_key in self.qty_dict.keys():
            if self.qty_dict[each_key] == 0 or self.qty_dict[each_key] == '':
                self.qty_dict.pop(each_key)


def buildOutputBook(wsheet, obj_sample_qty):
    
    #-----------------------------------------------------------------#
    #                         Initialization                          #
    #-----------------------------------------------------------------#
    col_name = 0
    col_item = 1
    col_qty = 2
    col_pcbvend = 3
    col_nor_eva = 4
    col_remark = 5
    col_dest_01 = 6
    col_dest_02 = 7
    col_dest_03 = 8

    l_value_title = {}
    l_value_title.setdefault(0, '')                  #0
    l_value_title.setdefault(1, 'Item')    #1 item
    l_value_title.setdefault(2, 'Quantity')    #2 qty
    l_value_title.setdefault(3, 'PCB Vendor')    #3 pcbvend
    l_value_title.setdefault(4, 'normal/eval')    #4 nor_eva
    l_value_title.setdefault(5, 'Remark')    #5 remark
    l_value_title.setdefault(6, 'Destination_01')    #6 dest_01
    l_value_title.setdefault(7, 'Destination_02')    #7 dest_02
    l_value_title.setdefault(8, 'Destination_03')    #8 dest_03

    nrow_starting = 1
    roll = 1
    nrow_normal = []
    req_qty_list = []
    real_qty_list = []

    #-----------------------------------------------------------------#
    #                 format setting for worksheet                    #
    #-----------------------------------------------------------------#
    # set format: alignment
    alignment = xlwt.Alignment()
    alignment.horz = xlwt.Alignment.HORZ_CENTER
    alignment.vert = xlwt.Alignment.VERT_CENTER
    # set format: pattern
    pattern = xlwt.Pattern()
    pattern.pattern = xlwt.Pattern.SOLID_PATTERN
    pattern.pattern_fore_colour = 5 #highlight in yellow
    # set format: borders
    borders = xlwt.Borders()
    borders.bottom = 1
    borders.top = 1
    borders.left = 1
    borders.right = 1
    # create style obj and cast above format setting
    style_gen = xlwt.XFStyle()
    style_gen.alignment = alignment
    style_gen.borders = borders
    # create a new style for build name cell only, from deepcoye based on style_gen
    style_highlight = copy.deepcopy(style_gen)
    style_highlight.pattern = pattern

    # write variant name
    wsheet.write(0, 0, obj_sample_qty.getVariant())
    # write build name
    wsheet.write(nrow_starting, 0, obj_sample_qty.getBuildName(), style_highlight)
    wsheet.col(0).width = 0x0d00    # set col width

    # write titles in one row, from col 01 to col 08
    for l in range(col_item, col_dest_03 + 1):
        wsheet.col(l + 3).width = 0x2000
        wsheet.write(nrow_starting, l, l_value_title[l], style_gen)

    # QuantitySample.qty_dict() order:
    # vcol_Normal(if any); vcol_Eval(if any); vcol_Normal_2nd(if any); vcol_Eval_2nd(if any); 
    # vcol_ModulePCB(if any); vcol_EvalPCB(if any); vcol_EvalPCBOnly(if any)

    # remove the empty/zero key-value pair from qty_dict
    obj_sample_qty.shrinkQtyDict()

    #-----------------------------------------------------------------#
    #     proc key-pair in qty_dict: key:item name/value:pcs qty      #
    #-----------------------------------------------------------------#
    for each_key in obj_sample_qty.qty_dict.keys():

    	print each_key

        pcs_factor = 0
        real_qty = 0        
        req_qty = int(obj_sample_qty.getQty(each_key))

        wsheet.write(nrow_starting + roll, col_item, roll, style_gen)
        wsheet.write(nrow_starting + roll, col_pcbvend, default_pcbvender, style_gen)
        wsheet.write(nrow_starting + roll, col_nor_eva, each_key, style_gen)

        # write remark
        if 'normal board' in str.lower(each_key):
            # for 'normal board', pcs qty must be multiple times of 16
            pcs_factor = MORNAL_ALIGN_QUAN        
            #nrow_normal = nrow_starting + roll
            nrow_normal.append(nrow_starting + roll)

            if (req_qty % pcs_factor != 0):
                real_qty = (int(req_qty / pcs_factor) + 1) * pcs_factor
                
                if (real_qty - req_qty) < EVAL_ALIGN_QUAN:
                    real_qty += pcs_factor
                #req_qty_list.append((int(req_qty / pcs_factor) + 1) * pcs_factor) # roll - 1
            else:
                real_qty = req_qty + pcs_factor
                #req_qty_list.append(req_qty + pcs_factor)  # roll - 1

        elif 'eval board' in str.lower(each_key):
            pcs_factor = EVAL_ALIGN_QUAN
            wsheet.write(nrow_starting + roll, col_remark, default_remark_eval, style_gen)

            if (req_qty % pcs_factor != 0):
                real_qty = (int(req_qty / pcs_factor) + 1) * pcs_factor
                #req_qty_list.append((int(req_qty / pcs_factor) + 1) * pcs_factor) # roll - 1
            else:
                real_qty = req_qty + pcs_factor
                #req_qty_list.append(req_qty + pcs_factor)  # roll - 1

        '''
        elif 'eval' in str.lower(each_key) and 'without module' not in str.lower(each_key) :
            # for 'eval board', pcs qty must be multiple times of 4
            pcs_factor = 4
            #default_remark = default_remark_eval     
            wsheet.write(nrow_starting + roll, col_remark, default_remark_eval, style_gen)

        elif 'eval' in str.lower(each_key) and 'without module' in str.lower(each_key) :
            # for 'eval board', pcs qty must be multiple times of 4
            pcs_factor = 4
            #default_remark = default_remark_eval     
            wsheet.write(nrow_starting + roll, col_remark, default_remark_eval, style_gen)
        '''

        # build and write real qty base on request qty and factor, for normal/eval
        print pcs_factor
        if pcs_factor != 0:
            req_qty_list.append(req_qty)
            real_qty_list.append(real_qty)
            
        wsheet.write(nrow_starting + roll, col_qty, str(int(req_qty)) + ' / ' + str(int(real_qty)), style_gen)
        
        # write destination with qty
        wsheet.write(nrow_starting + roll, col_dest_01, 'BJ: ' + str(int(req_qty)) + ' / ' + str(int(real_qty)), style_gen)
        wsheet.write(nrow_starting + roll, col_dest_02, 'BLN: 0 / 0', style_gen)
        wsheet.write(nrow_starting + roll, col_dest_03, 'DL: 0 / 0', style_gen)
        # set style
        wsheet.write(nrow_starting + roll, 0, '', style_gen) 

        roll += 1

    # write remark for normal
    '''
    if len(req_qty_list) != 0 and len(real_qty_list) != 0:            
        remark_normal = str(abs(req_qty_list[0] - req_qty_list[1])) + ' / ' + str(abs(real_qty_list[0] - real_qty_list[1])) + default_remark_normal
        wsheet.write(nrow_normal[0], col_remark, remark_normal, style_gen)
    '''
    if len(req_qty_list) != 0 and len(real_qty_list) != 0:
        for loop in range(0, len(req_qty_list), 2):           
            remark_normal = str(abs(req_qty_list[loop] - req_qty_list[loop + 1])) + ' / ' + str(abs(real_qty_list[loop] - real_qty_list[loop + 1])) + default_remark_normal
            wsheet.write(nrow_normal[loop], col_remark, remark_normal, style_gen)
    # if need to apend 2nd source quantity
    # for PCB, no need 3rd source


def analyInputTable(table, obj_sample_qty):
    
    vcol_Normal = ''
    vcol_Eval = ''
    vcol_Normal_2nd = ''
    vcol_Eval_2nd = ''
    vcol_ModulePCB = ''
    vcol_EvalPCB = ''
    vcol_EvalPCBOnly = ''
    
    try:

        obj_sample_qty.setBuildName(table.name)

        #print table.nrows
        #print table.ncols

        for nrow in range(table.nrows):
            for ncol in range(table.ncols):

                if 'planned' in str(table.cell_value(nrow, ncol)):
                    
                    # 1st line
                    str_value_cell = str(table.cell_value(0, ncol))

                    # 2nd line, for main source or second source
                    if 'second source' in str.lower(str(table.cell_value(1, ncol))):
                        # process for second source
                        if 'normal board' in str.lower(str(str_value_cell)):
                            vcol_Normal_2nd = str_value_cell
                            ncol_Normal_2nd_quan = ncol

                        elif 'eval board' in str.lower(str(str_value_cell)):
                            if vcol_Eval_2nd == '':
                                vcol_Eval_2nd = str_value_cell
                                ncol_Eval_2nd_quan = ncol
                        continue

                    # process for main source
                    if 'normal board' in str.lower(str(str_value_cell)):
                        vcol_Normal = str_value_cell
                        ncol_Normal_quan = ncol

                    elif 'eval board' in str.lower(str(str_value_cell)):
                    	if vcol_Eval == '':
	                        vcol_Eval = str_value_cell
	                        ncol_Eval_quan = ncol
                        elif 'without module' in str_value_cell:
                        	vcol_EvalPCBOnly = str_value_cell
                        	ncol_EvalPCBOnly_quan = ncol

                    elif 'Module PCB' in str_value_cell:
                        vcol_ModulePCB = str_value_cell
                        ncol_ModulePCB_quan = ncol

                    elif 'Eval PCB' in str_value_cell:
                    	vcol_EvalPCB = str_value_cell
                    	ncol_EvalPCB_quan = ncol

                if 'person in charge:' in str(table.cell_value(nrow, ncol)):
                    nrow_quan = nrow
                    if 0 != ncol:
                        print 'Warning for column num of \'person in charge\''
                    break

        # insert item name(key) and quantity(value) into class QuantitySample.qty_dict
        if vcol_Normal != '':
            obj_sample_qty.setQty(vcol_Normal, table.cell_value(nrow_quan, ncol_Normal_quan) + table.cell_value(nrow_quan, ncol_Eval_quan))
        if vcol_Eval != '':
            obj_sample_qty.setQty(vcol_Eval, table.cell_value(nrow_quan, ncol_Eval_quan))

        if vcol_Normal_2nd != '':
            obj_sample_qty.setQty(vcol_Normal, table.cell_value(nrow_quan, ncol_Normal_2nd_quan) + table.cell_value(nrow_quan, ncol_Eval_2nd_quan))
        if vcol_Eval_2nd != '':
            obj_sample_qty.setQty(vcol_Eval, table.cell_value(nrow_quan, ncol_Eval_2nd_quan))

        if vcol_ModulePCB != '':
            obj_sample_qty.setQty(vcol_ModulePCB, table.cell_value(nrow_quan, ncol_ModulePCB_quan))
        if vcol_EvalPCB != '':
            obj_sample_qty.setQty(vcol_EvalPCB, table.cell_value(nrow_quan, ncol_EvalPCB_quan))
        if vcol_EvalPCBOnly != '':
        	obj_sample_qty.setQty(vcol_EvalPCBOnly, table.cell_value(nrow_quan, ncol_EvalPCBOnly_quan))
                
    except Exception as e:
        print e
    finally:
        pass


def procExcel(in_fexcel, out_fname, variant_name):
       
    input_sheet = list('invalid')
    
    try:
        # access input book
        book_input = xlrd.open_workbook(in_fexcel)
        # create a new book as output
        book_sample_req = xlwt.Workbook()

        # scan input excel worksheet one by one
        for sheet_input in book_input.sheets():
            #sheet_input = input_book.sheet_by_index(2)

            # create obj for quantity collection
            obj_sample_qty = QuantitySample()

            obj_sample_qty.setVariant(variant_name)

            if 'A1' in sheet_input.name or 'A2' in sheet_input.name or 'B1' in sheet_input.name or 'B2' in sheet_input.name:
                # access and process worksheets from input book
                analyInputTable(sheet_input, obj_sample_qty)
                # append worksheet into output book
                sheet_sample_req = book_sample_req.add_sheet(sheet_input.name)
                # build sample request worksheet
                buildOutputBook(sheet_sample_req, obj_sample_qty)
            else:
                continue
        
    except Exception as e:
        print e
    finally:
        book_sample_req.save(out_fname)
        return

def procParam():
    usage_string = """ NPI_ResultCheck.py Usage: 
    ATC in list file execution and tracking;
    -i <fname>, --infile=<fname>    specify the ATC list file
    -o <fname>, --outfile=<fname>   specifies the output file (Default: atc.log)
    -v <variant name>, --variant=<variant name>
    -
    -h, --help                      print this usage message
    """
    # get program arguments
    try:
        (opts, args) = getopt.getopt(sys.argv[1:], 'hi:o:v:', ['help', 'infile=', 'outfile=', 'variant='])
    except getopt.GetoptError:
        print usage_string
        sys.exit(2)

    for opt, val in opts:
        if opt in ["-h", "--help"]:
            print usage_string
            sys.exit()

    in_fname = ''
    in_path = ''
    out_fname = ''
    variant_name = ''

    # take into account the last -i option
    for opt, val in opts:
        if opt in ['-i','--infile']:
            in_fname = val
    # if path included in in_fname, get the input file name first
    in_path = in_fname
    if '\\' in in_fname:        
        while ('\\' in in_fname):
            in_fname = in_fname[in_fname.find('\\')+1:]

    # take into account the last -o option
    for opt, val in opts:
        if opt in ['-o', '--outfile']:
            out_fname = val
    if out_fname == '':
        out_fname = 'Sample_Request.xls'

    # take into account the last -o option
    for opt, val in opts:
        if opt in ['-v', '--variant']:
            variant_name = str.upper(val)
    if variant_name == '':
        print 'Error: pls input variant name followed by \'-v\''
        return

    procExcel(in_path, out_fname, variant_name)

   
def main():
    procParam()

if __name__ == '__main__':
    main()
    

