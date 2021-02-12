def get_kanri_list(table):
    kanri_list = ''
    for row_cnt in range(len(table)):
        if table[row_cnt][0] == '管理':
            kanri_list += ('　・[' + table[row_cnt][7] + '] ' + table[row_cnt][8] + '\n')
    return kanri_list

def get_other_list(table):
    other = ''
    categoly = [table[i][0] for i in range(2, len(table))]
    categoly = list(set(categoly))
    for ctgl in categoly:
        if ctgl != '管理':
            other += '■' + ctgl + '\n'
            for row_cnt in range(len(table)):
                if table[row_cnt][0] == ctgl:
                    other += '　・[' + table[row_cnt][7] + '] ' + table[row_cnt][8] + '\n'
    return other
