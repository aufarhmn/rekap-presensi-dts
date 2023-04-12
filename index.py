import pandas as pd

source_df = pd.read_excel('raw12april.xlsx')
dest_df = pd.read_excel('rekap.xlsx')

dest_df.set_index('RegistNum', inplace=True)

for i, row in source_df.iterrows():
    regist_num = row['RegistNum']
    pertemuan = row['Pertemuan']
    kehadiran = row['Kehadiran']
    
    if regist_num in dest_df.index:
        if pertemuan == 'Pertemuan ke-1 -  6 Maret 2023':
            dest_df.loc[regist_num, 'Pertemuan 1'] = kehadiran
        elif pertemuan == 'Pertemuan ke-2 - 8 Maret 2023':
            dest_df.loc[regist_num, 'Pertemuan 2'] = kehadiran
        elif pertemuan == 'Pertemuan ke-3 - 9 Maret 2023':
            dest_df.loc[regist_num, 'Pertemuan 3'] = kehadiran
        elif pertemuan == 'Pertemuan ke-4 - 10 Maret 2023':
            dest_df.loc[regist_num, 'Pertemuan 4'] = kehadiran
        elif pertemuan == 'Pertemuan ke-5 - 13 Maret 2023':
            dest_df.loc[regist_num, 'Pertemuan 5'] = kehadiran
        elif pertemuan == 'Pertemuan ke-6 - 15 Maret 2023':
            dest_df.loc[regist_num, 'Pertemuan 6'] = kehadiran
        elif pertemuan == 'Pertemuan ke-7 - 17 Maret 2023' or pertemuan == 'Pertemuan ke-7 - 19 Maret 2023':
            dest_df.loc[regist_num, 'Pertemuan 7'] = kehadiran
        elif pertemuan == 'Pertemuan ke-8 - 20 Maret 2023':
            dest_df.loc[regist_num, 'Pertemuan 8'] = kehadiran
        elif pertemuan == 'Pertemuan ke-9 - 24 Maret 2023' or pertemuan == 'Pertemuan ke-9 - 22 Maret 2023':
            dest_df.loc[regist_num, 'Pertemuan 9'] = kehadiran
        elif pertemuan == 'Pertemuan ke-10 - 27 Maret 2023':
            dest_df.loc[regist_num, 'Pertemuan 10'] = kehadiran
        elif pertemuan == 'Pertemuan ke-11 - 29 Maret 2023':
            dest_df.loc[regist_num, 'Pertemuan 11'] = kehadiran
        elif pertemuan == 'Pertemuan ke-12 - 31 Maret 2023': 
            dest_df.loc[regist_num, 'Pertemuan 12'] = kehadiran
        elif pertemuan == 'Pertemuan ke-13 - 3 April 2023':
            dest_df.loc[regist_num, 'Pertemuan 13'] = kehadiran
        elif pertemuan == 'Pertemuan ke-14 - 5 April 2023':
            dest_df.loc[regist_num, 'Pertemuan 14'] = kehadiran
        elif pertemuan == 'Pertemuan ke-15 - 7 April 2023' or pertemuan == 'Pertemuan ke-15 - 6 April 2023':
            dest_df.loc[regist_num, 'Pertemuan 15'] = kehadiran
        elif pertemuan == 'Pertemuan ke-16 - 12 April 2023':
            dest_df.loc[regist_num, 'Pertemuan 16'] = kehadiran
        elif pertemuan == 'Pertemuan ke-17 - 14 April 2023':
            dest_df.loc[regist_num, 'Pertemuan 17'] = kehadiran
        
dest_df.reset_index(inplace=True)

dest_df.to_excel('Rekap 12 April (10 PM).xlsx', index=False)