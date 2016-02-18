import os, glob
import openpyxl
from os.path import join
from openpyxl import load_workbook
import pandas as pd
from flat_tools import tbl_cols_order, heat, save_excel, expand_rows, scene_order_clean, strip_title
import datetime, time

from sqlalchemy import create_engine # database connection
disk_engine = create_engine('sqlite:///flattenizer.db')

pd.options.mode.chained_assignment = None  # default='warn'

dir = os.path.dirname(__file__)
desktop = join(os.path.expanduser('~'),'Desktop')


class Flattenizer:
    def __init__(self, target_co, flat_dir):
        self.target_co = target_co
        self.flat_dir = flat_dir

    def extract(self):
        extract_df = pd.read_excel(self.target_co,skiprows=[0,1,2], sheetname=0, index=False)#.astype(str).replace('nan','')#, usecols = self.usecols).rename(columns=self.rename)
        rename_dict = {key: value for (key, value) in zip(tbl_cols_order('extract_cols'), tbl_cols_order('extract_rename'))}
        extract_df = extract_df.ix[:, tbl_cols_order('extract_cols')].rename(columns=rename_dict)
        extract_df2 = extract_df.ix[:, tbl_cols_order('transform_cols')]#.rename(columns=rename_dict)

        try:
            for br in ['br_block_clean', 'br_scene_clean']:
                extract_df2[br] = extract_df2[br.replace('_clean','')].map(lambda x: 0 if pd.isnull(x) else x).astype(int)
                extract_df2[br] = extract_df2[br].apply(lambda x: str(x).zfill(7) if int(x) >619000 else str(x).zfill(6)).astype(str)
        except:
            pass

        def merge_id(row):
            concat = str(row['br_block_clean']) + '|' + str(row['br_scene_clean'])            
            concat = concat.replace('.0','')
            return concat


        extract_df2['merge_id'] = extract_df2.apply(merge_id, axis=1)
        self.out_filename = join(self.flat_dir,'FLAT_' + os.path.basename(self.target_co))
        
        extract_df2.to_excel(self.out_filename, 'extract', index=False)
        self.extract = extract_df2

    def transform(self):
        scene_flat_df1 = self.extract[self.extract['house_id'].notnull()]
        scene_flat_df1['heat'] = scene_flat_df1.apply(heat, axis=1).astype('str').replace('nan','') #tools
        scene_flat_df2 = expand_rows(scene_flat_df1,'heat',',') #tools
        scene_flat_df2['br'] = scene_flat_df2['br_scene_clean']
        scene_flat_df2['title'] = scene_flat_df2['scene_title']
        scene_flat_df2 = scene_flat_df2[scene_flat_df2['heat'] != 'nan']


        block_flat_df1 = scene_flat_df2[scene_flat_df2['block_heat'].notnull()]
        block_flat_df1['heat'] = block_flat_df1['block_heat']
        block_flat_df1 = expand_rows(block_flat_df1,'heat',',') #tools
        block_flat_df1['br'] = block_flat_df1['br_block_clean']
        block_flat_df1['title'] = block_flat_df1['block_title']        
        block_flat_df1['scene_status'] = block_flat_df1['merge_id'].map(lambda x: 'Block Only' if x.split('|')[1] == '000000' else 'BLOCK')
        

        '''where i need to fix to account for block only'''
        def block_merge_id(row):
            if str(row['merge_id']).split('|')[1] != '000000':
                return str(row['merge_id']).split('|')[0] + '|' + str(row['merge_id']).split('|')[0]
            else:
                return row['merge_id']


        block_flat_df1['merge_id'] = block_flat_df1.apply(block_merge_id, axis=1) 

        #block_flat_df1['merge_id'] = block_flat_df1['merge_id'].apply(lambda x: str(x).split('|')[0] + '|' + str(x).split('|')[0] + '|' + str(x).split('|')[2])


        block_flat_df2 = block_flat_df1.drop(['cast'], axis=1).reset_index()
        block_cast_df = block_flat_df1.ix[:, tbl_cols_order('block_cast')]
        block_cast_df = block_cast_df[block_cast_df['cast'].notnull()]
        block_cast_df = pd.DataFrame({'cast':block_cast_df.groupby(['merge_id', 'br_block_clean'])['cast'].apply(lambda x: "%s" % ', '.join(list(set(x))))}).reset_index()
        block_cast_df = block_cast_df[['br_block_clean','cast']]
        block_flat_df2 = pd.merge(block_flat_df2, block_cast_df, how='left', on='br_block_clean')#.drop('cast_x', axis=1)

        block_scene_concat_df1 = pd.concat([scene_flat_df2, block_flat_df2])
        block_scene_concat_df1 = block_scene_concat_df1[block_scene_concat_df1['title'].notnull()]
        
        block_scene_concat_df1['scene_order_clean'] = block_scene_concat_df1.apply(scene_order_clean, axis=1) #tools
        block_scene_concat_df1['language'] = block_scene_concat_df1['language'].apply(lambda x: 'en' if pd.isnull(x) else x)
        block_scene_concat_df1['house_id'] = block_scene_concat_df1.apply(lambda row: None if row['scene_status'] == 'BLOCK' else row['house_id'], axis=1) #I should have learned this a long ass time ago
        block_scene_concat_df1['scene_notes'] = block_scene_concat_df1.apply(lambda row: None if row['scene_status'] == 'BLOCK' else row['scene_notes'], axis=1) #I should have learned this a long ass time ago
        
        block_scene_concat_df1['block_title_stripped'] = block_scene_concat_df1.apply(lambda x: strip_title(x, 'block_title'), axis=1)
        block_scene_concat_df1['title_stripped'] = block_scene_concat_df1.apply(lambda x: strip_title(x, 'title'), axis=1)

        '''write something for TMOs and VO_DUB'''
        

        '''SORT COLUMNS'''
        index_group_df = block_scene_concat_df1[block_scene_concat_df1['br_block_clean'] != '000000']
        index_group_df = index_group_df[['br_block_clean', 'index']].rename(columns={'index':'index_sort'})
        index_group_df = index_group_df.groupby('br_block_clean')['index_sort'].idxmin() + 0#makes it match the content order index
        index_group_df = index_group_df.reset_index()
        
        block_scene_concat_df2 = pd.merge(block_scene_concat_df1, index_group_df, how='left', on='br_block_clean')

        def only_sort(row):
            if row['scene_status'] == 'Block Only':
                return 2

            if row['scene_order_clean'] == 'Scene Only':
                return 3

            else:
                return 1


        block_scene_concat_df2['scene_only_sort'] = block_scene_concat_df2.apply(only_sort, axis=1)

        #block_scene_concat_df2['scene_only_sort'] = block_scene_concat_df2.apply(lambda row: 2 if row['scene_order_clean'] == 'Scene Only' else 1, axis=1)
        #block_scene_concat_df2['scene_only_sort'] = block_scene_concat_df2.apply(lambda row: 3 if row['scene_order_clean'] == 'Block Only' else 1, axis=1)
        block_scene_concat_df2 = block_scene_concat_df2.sort_values(by=['scene_only_sort', 'index_sort','br_block_clean','heat','scene_order_clean',])
       # block_scene_concat_df2.to_excel('scene_only_test.xlsx')


        block_scene_concat_df2 = block_scene_concat_df2.drop_duplicates(subset=['merge_id', 'heat']).reset_index(drop=True)
        block_scene_concat_df2 = block_scene_concat_df2.ix[:, tbl_cols_order('stage_cols')]
        
        self.transform_df = block_scene_concat_df2
        save_excel(self.out_filename, 'transform', self.transform_df) #tools


    def load(self):
        load_df = self.transform_df.ix[:, tbl_cols_order('google_sheet')]

        #change scene_notes to edit_notes so that Elliots notes can transfer
        load_df = load_df.rename(columns={'scene_notes':'edit_notes'})
        load_df['br_block_clean'] = load_df['br_block_clean'].str.replace('000000','')
        load_df['date_requested/added'] = datetime.datetime.fromtimestamp(time.time()).strftime('%m/%d/%Y %H:%M:%S')        

        


        save_excel(self.out_filename, 'output', load_df) #tools

        print ('flattenize complete!')






#def flattenizer(co_template=r'C:\Users\A_DO\Dropbox\1. Projects\2. Python\Flattenizer_Pandas\1. Content Order\Shortform RFW.xlsx', flat_co='FLAT-' + 'Shortform RFW.xlsx'):
def flattenizer(target_co, flat_dir):
    #Creating the class attribute here
    co = Flattenizer(target_co, flat_dir)

    co.extract()
    co.transform()
    co.load()




#test_co = r'C:\Users\A_DO\Dropbox\Active Python\Flattenizer_Pandas2\practice COs\UKEUCO_2016_04.xlsx'
#test_co = '/Users/RawDawgAss/Dropbox/Active Python/Flattenizer_Pandas2/practice COs/UKEUCO_2016_04.xlsx'
#flattenizer(test_co, dir)