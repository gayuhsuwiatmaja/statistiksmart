# -*- coding: utf-8 -*-
"""
Created on Fri Nov 15 13:12:53 2019

@author: User
"""

import streamlit as st
import pandas as pd
import numpy as np
import time
import matplotlib.pyplot as plt
import matplotlib
matplotlib.use('Agg')

@st.cache(allow_output_mutation=True)
def gabung_data(namafile):
    nama_file = str(namafile)
    xl = pd.ExcelFile(nama_file)
    namaSheet = xl.sheet_names
    
    hasil = []
    for i in range(len(namaSheet)):
        hasil.append(pd.read_excel(nama_file,namaSheet[i]))
    #melakukan penambahan dataframe jadi hasilnya adalah dataframe yang merupakan penggabungan semua sheet yang ada nama bulannya
    result = pd.concat(hasil)
    return result

def baca_data(namafile2):
    nama_file2 = str(namafile2)
    xl2 = pd.ExcelFile(nama_file2)
    namaSheet2 = xl2.sheet_names
    namaSheet2.append('Semua Bulan')
    return namaSheet2

#ukuran huruf yang ada di pie chart
plt.rcParams['font.size'] = 20

# xl = pd.ExcelFile('bulan/STATISTIK_PORTAL_EOFFICE_APRIL_NOV 2019.xlsx')
# namaSheet = xl.sheet_names
# namaSheet.append('Semua Bulan')

namaSheet3 = baca_data('bulan/STATISTIK_PORTAL_EOFFICE_APRIL_NOV 2019.xlsx')
option3 = st.selectbox('Bulan', namaSheet3)

if option3 == "Semua Bulan":
    
    # ambil semua nama sheet bulan tanpa pilihan "Semua"
    # namaSheet2 = xl.sheet_names
    
    # melakukan penambahan data berformat dataframe dari semua sheet yang ada nama bulannya
    # hasil = []
    # for i in range(len(namaSheet2)):
        # hasil.append(pd.read_excel('STATISTIK_PORTAL_EOFFICE_APRIL_NOV 2019.xlsx',namaSheet2[i]))
    # melakukan penambahan dataframe jadi hasilnya adalah dataframe yang merupakan penggabungan semua sheet yang ada nama bulannya
    # result = pd.concat(hasil)
    
    result = gabung_data('bulan/STATISTIK_PORTAL_EOFFICE_APRIL_NOV 2019.xlsx')
    #tmbah nilai "Semua" di pilihan Eselon 1
    TmbahSemuaEs1 = result.append({'eselon_1' : 'Semua'} , ignore_index=True)
    
    #tmbah pilihan "semua" di eselon 2
    #TmbahSemuaEs2 = df2.append({'eselon_2' : 'Semua'} , ignore_index=True)
    #st.write(TmbahSemuaEs1['eselon_1'].unique())
    option1 = st.selectbox('Pilih es1', TmbahSemuaEs1['eselon_1'].unique())
    #st.write(option1)
    #convert string ke list
    listOp1 = [option1]
    #menampilkan eselon 2 kebawah yang es 1 nya berasal dari option1
    rslt_Op1 = result[result['eselon_1'].isin(listOp1)]
    
    #tmbah nilai "Semua" di pilihan Eselon 2
    TmbahSemuaEs2 = rslt_Op1.append({'eselon_2' : 'Semua'} , ignore_index=True)
    HasilEs2 = TmbahSemuaEs2['eselon_2']
    
    
    
    option2 = st.selectbox('Pilih es2', HasilEs2.unique())
    
    
    blankIndex=[''] * len(result)
    result.index=blankIndex
    
    es1 = [option1]
    es2 = [option2] 
    
    if option1 == "Semua":  
        rslt_TdkAkses = result[(result['login_portal'] == 0)]
    
        rslt_Akses = result[(result['login_portal'] > 0)]
    
        jmlTidakAkses = len(rslt_TdkAkses)
        jmlAkses = len(rslt_Akses)
    
    
        # Pie chart, where the slices will be ordered and plotted counter-clockwise:
        labels = 'Akses Smart', 'Tidak Akses Smart', 
        sizes = [jmlAkses,jmlTidakAkses]
        colors = ['blue', 'red']
        explode = (0.1, 0)  # only "explode" the 2nd slice (i.e. 'Hogs')
    
        fig1, ax1 = plt.subplots()
        ax1.pie(sizes, explode=explode, colors=colors, autopct='%1.1f%%',shadow=True, startangle=90)
        ax1.legend(labels,loc='upper center', bbox_to_anchor=(0.5, -0.05),fancybox=True, shadow=True, ncol=5, fontsize = 11)
        ax1.axis('equal')  # Equal aspect ratio ensures that pie is drawn as a circle.
    
        #plt.show()
        
        plt.title(option3, bbox={'facecolor':'0.8', 'pad':5})
        st.pyplot(plt)

        
        #jumlah kan nama yg masuk dlm kriteria ini
        TotalNama = [rslt_Akses['nama'],rslt_TdkAkses['nama']]
        hasil = pd.concat(TotalNama)
        
        #nama2 yang akses berdasarkan kriteria ini
        namaAkses = rslt_Akses['nama']
        #nama tidak akses adalah nama2 yang ada di total nama yang tidak tercantum di nama akses
        nama_TdkAkses = hasil[~hasil.isin(namaAkses) ]
        
        
        #menampilkan pegawai dengan kriteria tertentu
        if st.checkbox('Tunjukkan siapa saja yang tidak akses SMART'):
            st.write(nama_TdkAkses.unique())
        
        if st.checkbox('Tunjukkan siapa saja yang akses SMART'):
            #st.write(rslt_Akses[['nama','login_portal']])
            st.write(rslt_Akses['nama'].unique())
       
    else:
        if option2 == "Semua":
            rslt_TdkAkses = result[(result['login_portal'] == 0) & 
                             result['eselon_1'].isin(es1)]
        
            rslt_Akses = result[(result['login_portal'] > 0) & 
                             result['eselon_1'].isin(es1)]
        
            jmlTidakAkses = len(rslt_TdkAkses)
            jmlAkses = len(rslt_Akses)
        
        
            # Pie chart, where the slices will be ordered and plotted counter-clockwise:
            labels = 'Akses Smart', 'Tidak Akses Smart', 
            sizes = [jmlAkses,jmlTidakAkses]
            colors = ['blue', 'red']
            explode = (0.1, 0)  # only "explode" the 2nd slice (i.e. 'Hogs')
        
            fig1, ax1 = plt.subplots()
            ax1.pie(sizes, explode=explode, colors=colors, autopct='%1.1f%%',shadow=True, startangle=90)
            ax1.legend(labels,loc='upper center', bbox_to_anchor=(0.5, -0.05),fancybox=True, shadow=True, ncol=5, fontsize = 11)
            ax1.axis('equal')  # Equal aspect ratio ensures that pie is drawn as a circle.
        
            #plt.show()
            
            plt.title(option3, bbox={'facecolor':'0.8', 'pad':5})
            st.pyplot(plt)
        
            #jumlah kan nama yg masuk dlm kriteria ini
            TotalNama = [rslt_Akses['nama'],rslt_TdkAkses['nama']]
            hasil = pd.concat(TotalNama)
            
            #nama2 yang akses berdasarkan kriteria ini
            namaAkses = rslt_Akses['nama']
            #nama tidak akses adalah nama2 yang ada di total nama yang tidak tercantum di nama akses
            nama_TdkAkses = hasil[~hasil.isin(namaAkses) ]
            
            #menampilkan pegawai dengan kriteria tertentu
            if st.checkbox('Tunjukkan siapa saja yang tidak akses SMART'):
                st.write(nama_TdkAkses.unique())
            
            if st.checkbox('Tunjukkan siapa saja yang akses SMART'):
                #st.write(rslt_Akses[['nama','login_portal']])
                st.write(rslt_Akses['nama'].unique())
        else:
            #st.write("sukses")
            #hilangkan index paling kiri
            rslt_TdkAkses = result[(result['login_portal'] == 0) & 
                                result['eselon_1'].isin(es1) &
                                result['eselon_2'].isin(es2)]
        
            rslt_Akses = result[(result['login_portal'] > 0) & 
                             result['eselon_1'].isin(es1) &
                             result['eselon_2'].isin(es2)]
          
         
            jmlTidakAkses = len(rslt_TdkAkses)
            jmlAkses = len(rslt_Akses)
        
        
            # Pie chart, where the slices will be ordered and plotted counter-clockwise:
            labels = 'Akses Smart', 'Tidak Akses Smart', 
            sizes = [jmlAkses,jmlTidakAkses]
            colors = ['blue', 'red']
            explode = (0.1, 0)  # only "explode" the 2nd slice (i.e. 'Hogs')
        
            fig1, ax1 = plt.subplots()
            ax1.pie(sizes, explode=explode, colors=colors, autopct='%1.1f%%',shadow=True, startangle=90)
            #ax1.legend(labels, loc="best", fontsize=15,bbox_transform=plt.gcf().transFigure)
            ax1.legend(labels,loc='upper center', bbox_to_anchor=(0.5, -0.05),fancybox=True, shadow=True, ncol=5, fontsize = 11)
            ax1.axis('equal')  # Equal aspect ratio ensures that pie is drawn as a circle.
        
            #plt.show()
            
            plt.title(option3, bbox={'facecolor':'0.8', 'pad':5})
            st.pyplot(plt)
        
            #jumlah kan nama yg masuk dlm kriteria ini
            TotalNama = [rslt_Akses['nama'],rslt_TdkAkses['nama']]
            hasil = pd.concat(TotalNama)
            
            #nama2 yang akses berdasarkan kriteria ini
            namaAkses = rslt_Akses['nama']
            #nama tidak akses adalah nama2 yang ada di total nama yang tidak tercantum di nama akses
            nama_TdkAkses = hasil[~hasil.isin(namaAkses) ]
            
            #menampilkan pegawai dengan kriteria tertentu
            if st.checkbox('Tunjukkan siapa saja yang tidak akses SMART'):
                st.write(nama_TdkAkses.unique())
                
            if st.checkbox('Tunjukkan siapa saja yang akses SMART'):
                #st.write(rslt_Akses[['nama','login_portal']])
                st.write(rslt_Akses['nama'].unique())
        
else:
    #ambil nilai dari satu file excel dan satu sheet tertentu
    df2 = pd.read_excel('bulan/STATISTIK_PORTAL_EOFFICE_APRIL_NOV 2019.xlsx',option3)
    #tmbah nilai "Semua" di pilihan Eselon 1
    TmbahSemuaEs1 = df2.append({'eselon_1' : 'Semua'} , ignore_index=True)
    #tmbah pilihan "semua" di eselon 2
    #TmbahSemuaEs2 = df2.append({'eselon_2' : 'Semua'} , ignore_index=True)
    #st.write(TmbahSemuaEs1['eselon_1'].unique())
    option1 = st.selectbox('Pilih es1', TmbahSemuaEs1['eselon_1'].unique())
    #st.write(option1)
    #convert string ke list
    listOp1 = [option1]
    #menampilkan eselon 2 kebawah yang es 1 nya berasal dari option1
    rslt_Op1 = df2[df2['eselon_1'].isin(listOp1)]
    
    #tmbah nilai "Semua" di pilihan Eselon 2
    TmbahSemuaEs2 = rslt_Op1.append({'eselon_2' : 'Semua'} , ignore_index=True)
    HasilEs2 = TmbahSemuaEs2['eselon_2']
    
    option2 = st.selectbox('Pilih es2', HasilEs2.unique())
    
    blankIndex=[''] * len(df2)
    df2.index=blankIndex
    
    es1 = [option1]
    es2 = [option2]
    if option1 == "Semua":  
        rslt_TdkAkses = df2[(df2['login_portal'] == 0)]
    
        rslt_Akses = df2[(df2['login_portal'] > 0)]
    
        jmlTidakAkses = len(rslt_TdkAkses)
        jmlAkses = len(rslt_Akses)
    
    
        # Pie chart, where the slices will be ordered and plotted counter-clockwise:
        labels = 'Akses Smart', 'Tidak Akses Smart', 
        sizes = [jmlAkses,jmlTidakAkses]
        colors = ['blue', 'red']
        explode = (0.1, 0)  # only "explode" the 2nd slice (i.e. 'Hogs')
    
        fig1, ax1 = plt.subplots()
        ax1.pie(sizes, explode=explode, colors=colors,autopct='%1.1f%%',shadow=True, startangle=90)
        ax1.legend(labels,loc='upper center', bbox_to_anchor=(0.5, -0.05),fancybox=True, shadow=True, ncol=5, fontsize = 11)
        ax1.axis('equal')  # Equal aspect ratio ensures that pie is drawn as a circle.
    
        #plt.show()
        
        plt.title(option3, bbox={'facecolor':'0.8', 'pad':5})
        st.pyplot(plt)
    
        #menampilkan pegawai dengan kriteria tertentu
        if st.checkbox('Tunjukkan siapa saja yang tidak akses SMART'):
            st.write(rslt_TdkAkses['nama'].unique())
        
        if st.checkbox('Tunjukkan siapa saja yang akses SMART'):
            #st.write(rslt_Akses[['nama','login_portal']])
            st.write(rslt_Akses['nama'].unique())
       
    else:
        if option2 == "Semua":
            rslt_TdkAkses = df2[(df2['login_portal'] == 0) & 
                             df2['eselon_1'].isin(es1)]
        
            rslt_Akses = df2[(df2['login_portal'] > 0) & 
                             df2['eselon_1'].isin(es1)]
        
            jmlTidakAkses = len(rslt_TdkAkses)
            jmlAkses = len(rslt_Akses)
        
        
            # Pie chart, where the slices will be ordered and plotted counter-clockwise:
            labels = 'Akses Smart', 'Tidak Akses Smart', 
            sizes = [jmlAkses,jmlTidakAkses]
            colors = ['blue', 'red']
            explode = (0.1, 0)  # only "explode" the 2nd slice (i.e. 'Hogs')
        
            fig1, ax1 = plt.subplots()
            ax1.pie(sizes, explode=explode, colors=colors, autopct='%1.1f%%',shadow=True, startangle=90)
            ax1.legend(labels,loc='upper center', bbox_to_anchor=(0.5, -0.05),fancybox=True, shadow=True, ncol=5, fontsize = 11)
            ax1.axis('equal')  # Equal aspect ratio ensures that pie is drawn as a circle.
        
            #plt.show()
            
            plt.title(option3, bbox={'facecolor':'0.8', 'pad':5})
            st.pyplot(plt)
        
            #menampilkan pegawai dengan kriteria tertentu
            if st.checkbox('Tunjukkan siapa saja yang tidak akses SMART'):
                st.write(rslt_TdkAkses['nama'].unique())
            
            if st.checkbox('Tunjukkan siapa saja yang akses SMART'):
                #st.write(rslt_Akses[['nama','login_portal']])
                st.write(rslt_Akses['nama'].unique())
        else:
            #st.write("sukses")
            #hilangkan index paling kiri
            rslt_TdkAkses = df2[(df2['login_portal'] == 0) & 
                                df2['eselon_1'].isin(es1) &
                                df2['eselon_2'].isin(es2)]
        
            rslt_Akses = df2[(df2['login_portal'] > 0) & 
                             df2['eselon_1'].isin(es1) &
                             df2['eselon_2'].isin(es2)]
          
         
            jmlTidakAkses = len(rslt_TdkAkses)
            jmlAkses = len(rslt_Akses)
        
        
            # Pie chart, where the slices will be ordered and plotted counter-clockwise:
            labels = 'Akses Smart', 'Tidak Akses Smart', 
            sizes = [jmlAkses,jmlTidakAkses]
            colors = ['blue', 'red']
            explode = (0.1, 0)  # only "explode" the 2nd slice (i.e. 'Hogs')
        
            fig1, ax1 = plt.subplots()
            ax1.pie(sizes, explode=explode, colors=colors,autopct='%1.1f%%',shadow=True, startangle=90)
            #ax1.legend(labels, loc="best", fontsize=15,bbox_transform=plt.gcf().transFigure)
            ax1.legend(labels,loc='upper center', bbox_to_anchor=(0.5, -0.05),fancybox=True, shadow=True, ncol=5, fontsize = 11)
            ax1.axis('equal')  # Equal aspect ratio ensures that pie is drawn as a circle.
        
            #plt.show()
            
            plt.title(option3, bbox={'facecolor':'0.8', 'pad':5})
            st.pyplot(plt)
        
            #menampilkan pegawai dengan kriteria tertentu
            if st.checkbox('Tunjukkan siapa saja yang tidak akses SMART'):
                st.write(rslt_TdkAkses['nama'].unique())   
                
            if st.checkbox('Tunjukkan siapa saja yang akses SMART'):
                #st.write(rslt_Akses[['nama','login_portal']])
                st.write(rslt_Akses['nama'].unique())
        
