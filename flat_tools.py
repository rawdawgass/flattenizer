import pandas as pd
from os.path import join
from openpyxl import load_workbook
import re




def tbl_cols_order(usecol):
    cols_df = pd.read_excel(join('tables', 'flat_cols.xlsx'))
    return cols_df[cols_df[usecol].notnull()][usecol].tolist()


def save_excel(output_dir, tab, df):
    book = load_workbook(output_dir)
    writer = pd.ExcelWriter(output_dir, engine='openpyxl', datetime_format='YYYY-m-d')
    writer.book=book
    df.to_excel(writer,tab, index=False)
    writer.save()





def heat(row): #this function puts the block heat in the scene heat column so that we can flatten it later 
    block_heats = str(row['block_heat']).split(',')
    scene_heats = str(row['scene_heat']).split(',')
    all_heats = [set(block_heats + scene_heats)]
    all_heats = str([",".join(map(str,x)) for x in all_heats]).strip('[]').strip("''")
    return all_heats

def expand_rows(df,target_column,separator): #miracle function
    # df = dataframe to split,
    #target_column = the column containing the values to split
    #separator = the symbol used to perform the split

    #returns: a dataframe with each entry for the target column separated, with each element moved into a new row. 
    #The values in the other columns are duplicated across the newly divided rows.
    
    def splitListToRows(row,row_accumulator,target_column,separator):
        split_row = row[target_column].split(separator)
        for s in split_row:
            new_row = row.to_dict()
            new_row[target_column] = s
            row_accumulator.append(new_row)
    new_rows = []
    df.apply(splitListToRows,axis=1,args = (new_rows,target_column,separator))
    new_df = pd.DataFrame(new_rows)
    return new_df



def strip_title(row, title_col): #figure out how to combine the two functions later
    try:
        replace = row[title_col].replace('#','num').replace('+','plus').replace('"','_inch').replace('&','and') #replace character strings here
        clean = re.sub('[^0-9a-zA-Z_\s]+', '', replace.lower())
        clean =  ' '.join(clean.split())
        final_clean = clean.replace(' ','_')
        return final_clean
    except:
        #raise
        pass



def scene_order_clean(row):
    if 'block' in row['scene_status'].lower():
        return ''
    if row['merge_id'].split('|')[0] == '000000':
        return 'Scene Only'
    
    if row['heat'] not in str(row['block_heat']).split(','):
        return 'Scene Only'
    
    else:
        return row['scene_order']