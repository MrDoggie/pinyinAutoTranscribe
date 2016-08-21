from openpyxl import load_workbook
from openpyxl import Workbook
import re

wb = load_workbook (filename = 'data.xlsx')

ws = wb.active

rows = ws.rows # get row numbers 

#c = ws.cell(row = 2, column = 1).value

#print c

# ws.cell (row=2, column = 2).value = 'ch'
# ws.cell (row=2, column = 3).value = '0'
# ws.cell (row=2, column = 4).value = 'r'
# ws.cell (row=2, column = 5).value = '0'
# ws.cell (row=2, column = 6).value = 'i'



#fuyin options: bpmfdtnlkghjqxzcsr; have rule out impossible combination

for rn in range(len(rows)):
	traditionInput=str(ws.cell(row=rn+1, column = 1).value)
	m_i1 = re.search('^([zhcsrz]+)i$' , traditionInput) # -i after zh, ch, sh,r,z,c,s
	m_i2 = re.search('^([bpmdtnljqxy])i$',traditionInput) #-i after 0, b, p,m,d,t,n,j,q,x; y will be handled later 
	m_a = re.search('^([bpmfdtnlkhzcs]+)a$',traditionInput) #-a
	m_o = re.search('^([bpmfl]+)o$', traditionInput) #-o
	m_e = re.search('^([mdtnlkghzcsr]+)e$',traditionInput) #-e
	m_ai = re.search('^([bpmdtnlkghzcs]+)ai$', traditionInput) #-ai
	m_ei = re.search('^([bpmfdtnlkghsz]+)ei$',traditionInput) #-ei
	m_ao = re.search('^([bpmdtnlkghzcsr]+)ao$',traditionInput) #-ao
	m_ou = re.search('^([pmfdtlkghzcs]+)ou$',traditionInput) #-ou
	m_an = re.search('^([bpmfdtnlkghzcsr]+)an$',traditionInput) #-an
	m_en = re.search('^([bpmfdnkghzcsr]+)en$',traditionInput)#-an
	m_ang = re.search('^([bpmfdtnlkghzcsr]+)ang$', traditionInput) #-ang
	m_eng = re.search('^([bpmfdtnlkghzcsr]+)eng$',traditionInput) #-eng
	m_er = re.search('^er$', traditionInput) #er, no other possibility
	m_ia = re.search('^([dljqx])ia$', traditionInput) #-ia ** TO DO y special case
	m_ie = re.search('^([bpmdtnljqx])ie$', traditionInput) #-ie 
	m_iu = re.search('^([mdnljqx])iu$',traditionInput) #-iu
	m_ian = re.search('^([bpmdtnljqx])ian$',traditionInput) #-ian
	m_in = re.search('^([bpmnljqxy])in$',traditionInput) #-in; y will be handled in the swtich branch
	m_iang = re.search('^([bnljqx])iang$', traditionInput) #-iang
	m_ing = re.search('^([bpmdtnljqxy])ing$', traditionInput) #-ing; y will be handled in the switch branch
	m_u = re.search('^([bpmfdtnlkghzcsrw]+)u$', traditionInput) #-ing (ju qu xu yu are exception); wu will be handled in the swtich branch
	m_ua = re.search('^([kghsz]+)ua$',traditionInput) #-ua
	m_uo = re.search('^([dtnlkghzcsr]+)uo$',traditionInput) #-uo
	m_uai = re.search('^([kghzcs]+)uai$',traditionInput) #-uai
	m_ui = re.search('^([dtkghzcsr]+)ui$',traditionInput) #-ui
	m_uan = re.search('^([dtnlkghzcsr]+)uan$',traditionInput) #-uan
	m_un = re.search('^([dtnlkghzcsr]+)un$',traditionInput) #-un
	m_uang = re.search('^([kghzcs]+)uang$',traditionInput) #-uang
	m_ong = re.search('^([dtnlkghzcsr]+)ong$',traditionInput) #-ong
	m_U = re.search('^([jqxy])u$',traditionInput) #yu sound after jqx; yu handled in the switch branch
	m_Ue = re.search('^([jqxy])ue$',traditionInput) #yu sound in ue; yue handled in the switch branch
	m_Uan = re.search('^([jqxy])uan$',traditionInput) #yu sound in uan; yuan handled in the switch branch
	m_Un = re.search('^([jqxy])un$',traditionInput) #yu sound in un; yun handled in the switch branch
	m_iong = re.search('^([jqx])iong$',traditionInput) #-iong
	m_iao = re.search('^([bpmdtnljqx])iao$', traditionInput) #-iao
	#special cases
	m_ya = re.search('^ya$', traditionInput) #ya special case
	m_ye = re.search('^ye$', traditionInput) #ye special case
	m_you = re.search('^you$', traditionInput) #you special case
	m_yan = re.search('^yan$', traditionInput) #yan SC
	m_yang = re.search('^yang$', traditionInput) #yang SC
	m_wa = re.search('^wa$', traditionInput) #wa SC
	m_wo = re.search('^wo$', traditionInput) #wo SC
	m_wai = re.search('^wai$', traditionInput) #wai SC
	m_wei = re.search('^wei$', traditionInput) #wei SC
	m_wan = re.search('^wan$', traditionInput) #wan SC
	m_wang = re.search('^wang$', traditionInput) #wang SC
	m_yong = re.search('^yong$',traditionInput) #yong SC
	m_yao = re.search('^yao$', traditionInput) #yao SC
	#switch branches
	if m_i1:
		ws.cell(row = rn+1, column = 2).value=m_i1.group(1) 
		ws.cell(row = rn+1, column = 3).value='0'
		ws.cell(row = rn+1, column = 4).value='r'
		ws.cell(row = rn+1, column = 5).value='0'
		ws.cell(row = rn+1, column = 6).value='i'
	elif m_i2:
		ws.cell(row = rn+1, column = 2).value=m_i2.group(1)
		if m_i2.group(1) == 'y':
			ws.cell(row = rn+1, column = 2).value = '0' #if input is yi, output 0 0 i 
		ws.cell(row = rn+1, column = 3).value='0'
		ws.cell(row = rn+1, column = 4).value='i'
		ws.cell(row = rn+1, column = 5).value='0'
		ws.cell(row = rn+1, column = 6).value='i'
	elif m_a:
		ws.cell(row = rn+1, column = 2).value=m_a.group(1) 
		ws.cell(row = rn+1, column = 3).value='0'
		ws.cell(row = rn+1, column = 4).value='a'
		ws.cell(row = rn+1, column = 5).value='0'
		ws.cell(row = rn+1, column = 6).value='a'
	elif m_o:
		ws.cell(row = rn+1, column = 2).value=m_o.group(1) 
		ws.cell(row = rn+1, column = 3).value='0'
		ws.cell(row = rn+1, column = 4).value='o'
		ws.cell(row = rn+1, column = 5).value='w'
		ws.cell(row = rn+1, column = 6).value='e'
	elif m_e:
		ws.cell(row = rn+1, column = 2).value=m_e.group(1) 
		ws.cell(row = rn+1, column = 3).value='0'
		ws.cell(row = rn+1, column = 4).value='e'
		ws.cell(row = rn+1, column = 5).value='0'
		ws.cell(row = rn+1, column = 6).value='e'
	elif m_ai:
		ws.cell(row = rn+1, column = 2).value=m_ai.group(1) 
		ws.cell(row = rn+1, column = 3).value='0'
		ws.cell(row = rn+1, column = 4).value='ai'
		ws.cell(row = rn+1, column = 5).value='0'
		ws.cell(row = rn+1, column = 6).value='ai'
	elif m_ei:
		ws.cell(row = rn+1, column = 2).value=m_ei.group(1) 
		ws.cell(row = rn+1, column = 3).value='0'
		ws.cell(row = rn+1, column = 4).value='ei'
		ws.cell(row = rn+1, column = 5).value='0'
		ws.cell(row = rn+1, column = 6).value='ei'
	elif m_ao:
		ws.cell(row = rn+1, column = 2).value=m_ao.group(1) 
		ws.cell(row = rn+1, column = 3).value='0'
		ws.cell(row = rn+1, column = 4).value='ao'
		ws.cell(row = rn+1, column = 5).value='0'
		ws.cell(row = rn+1, column = 6).value='au'
	elif m_ou:
		ws.cell(row = rn+1, column = 2).value=m_ou.group(1) 
		ws.cell(row = rn+1, column = 3).value='0'
		ws.cell(row = rn+1, column = 4).value='ou'
		ws.cell(row = rn+1, column = 5).value='0'
		ws.cell(row = rn+1, column = 6).value='eu'
	elif m_an:
		ws.cell(row = rn+1, column = 2).value=m_an.group(1) 
		ws.cell(row = rn+1, column = 3).value='0'
		ws.cell(row = rn+1, column = 4).value='an'
		ws.cell(row = rn+1, column = 5).value='0'
		ws.cell(row = rn+1, column = 6).value='an'
	elif m_en:
		ws.cell(row = rn+1, column = 2).value=m_en.group(1) 
		ws.cell(row = rn+1, column = 3).value='0'
		ws.cell(row = rn+1, column = 4).value='en'
		ws.cell(row = rn+1, column = 5).value='0'
		ws.cell(row = rn+1, column = 6).value='en'
	elif m_ang:
		ws.cell(row = rn+1, column = 2).value=m_ang.group(1) 
		ws.cell(row = rn+1, column = 3).value='0'
		ws.cell(row = rn+1, column = 4).value='ang'
		ws.cell(row = rn+1, column = 5).value='0'
		ws.cell(row = rn+1, column = 6).value='ang'
	elif m_eng:
		ws.cell(row = rn+1, column = 2).value=m_eng.group(1) 
		ws.cell(row = rn+1, column = 3).value='0'
		ws.cell(row = rn+1, column = 4).value='eng'
		ws.cell(row = rn+1, column = 5).value='0'
		ws.cell(row = rn+1, column = 6).value='eng'
	elif m_er: #only possible pinyin is er. Therefore no onset
		ws.cell(row = rn+1, column = 2).value='0'
		ws.cell(row = rn+1, column = 3).value='0'
		ws.cell(row = rn+1, column = 4).value='er'
		ws.cell(row = rn+1, column = 5).value='0'
		ws.cell(row = rn+1, column = 6).value='er'
	elif m_ia:
		ws.cell(row = rn+1, column = 2).value=m_ia.group(1) 
		ws.cell(row = rn+1, column = 3).value='i'
		ws.cell(row = rn+1, column = 4).value='a'
		ws.cell(row = rn+1, column = 5).value='i'
		ws.cell(row = rn+1, column = 6).value='a'
	elif m_ya: #deal with ya special case
		ws.cell(row = rn+1, column = 2).value='0' 
		ws.cell(row = rn+1, column = 3).value='i'
		ws.cell(row = rn+1, column = 4).value='a'
		ws.cell(row = rn+1, column = 5).value='i'
		ws.cell(row = rn+1, column = 6).value='a'
	elif m_ie: 
		ws.cell(row = rn+1, column = 2).value=m_ie.group(1) 
		ws.cell(row = rn+1, column = 3).value='i'
		ws.cell(row = rn+1, column = 4).value='e'
		ws.cell(row = rn+1, column = 5).value='i'
		ws.cell(row = rn+1, column = 6).value='e'
	elif m_ye: #ye special case
		ws.cell(row = rn+1, column = 2).value='0' 
		ws.cell(row = rn+1, column = 3).value='i'
		ws.cell(row = rn+1, column = 4).value='e'
		ws.cell(row = rn+1, column = 5).value='i'
		ws.cell(row = rn+1, column = 6).value='e'
	elif m_iu:
		ws.cell(row = rn+1, column = 2).value=m_iu.group(1) 
		ws.cell(row = rn+1, column = 3).value='i'
		ws.cell(row = rn+1, column = 4).value='ou'
		ws.cell(row = rn+1, column = 5).value='i'
		ws.cell(row = rn+1, column = 6).value='eu'
	elif m_you: #you special case
		ws.cell(row = rn+1, column = 2).value='0' 
		ws.cell(row = rn+1, column = 3).value='i'
		ws.cell(row = rn+1, column = 4).value='ou'
		ws.cell(row = rn+1, column = 5).value='i'
		ws.cell(row = rn+1, column = 6).value='eu'
	elif m_ian:
		ws.cell(row = rn+1, column = 2).value=m_ian.group(1) 
		ws.cell(row = rn+1, column = 3).value='i'
		ws.cell(row = rn+1, column = 4).value='an'
		ws.cell(row = rn+1, column = 5).value='i'
		ws.cell(row = rn+1, column = 6).value='an'
	elif m_yan:
		ws.cell(row = rn+1, column = 2).value='0' 
		ws.cell(row = rn+1, column = 3).value='i'
		ws.cell(row = rn+1, column = 4).value='an'
		ws.cell(row = rn+1, column = 5).value='i'
		ws.cell(row = rn+1, column = 6).value='an'
	elif m_in:
		ws.cell(row = rn+1, column = 2).value=m_in.group(1)
		if m_in.group(1) =='y': #handle yin special case
			ws.cell(row=rn+1, column = 2).value = '0' 
		ws.cell(row = rn+1, column = 3).value='0'
		ws.cell(row = rn+1, column = 4).value='in'
		ws.cell(row = rn+1, column = 5).value='0'
		ws.cell(row = rn+1, column = 6).value='in'
	elif m_iang:
		ws.cell(row = rn+1, column = 2).value=m_iang.group(1) 
		ws.cell(row = rn+1, column = 3).value='i'
		ws.cell(row = rn+1, column = 4).value='ang'
		ws.cell(row = rn+1, column = 5).value='i'
		ws.cell(row = rn+1, column = 6).value='ang'
	elif m_yang:
		ws.cell(row = rn+1, column = 2).value='0' 
		ws.cell(row = rn+1, column = 3).value='i'
		ws.cell(row = rn+1, column = 4).value='ang'
		ws.cell(row = rn+1, column = 5).value='i'
		ws.cell(row = rn+1, column = 6).value='ang'
	elif m_ing:
		ws.cell(row = rn+1, column = 2).value=m_ing.group(1) 
		if m_ing.group(1) == 'y': #ying  SC
			ws.cell(row =rn+1, column=2).value = '0'
		ws.cell(row = rn+1, column = 3).value='0'
		ws.cell(row = rn+1, column = 4).value='ing'
		ws.cell(row = rn+1, column = 5).value='i'
		ws.cell(row = rn+1, column = 6).value='eng'
	elif m_u:
		ws.cell(row = rn+1, column = 2).value=m_u.group(1)
		if m_u.group(1) =='w': #wu SC
			ws.cell(row=rn+1, column=2).value ='0' 
		ws.cell(row = rn+1, column = 3).value='0'
		ws.cell(row = rn+1, column = 4).value='u'
		ws.cell(row = rn+1, column = 5).value='0'
		ws.cell(row = rn+1, column = 6).value='u'
	elif m_ua:
		ws.cell(row = rn+1, column = 2).value=m_ua.group(1) 
		ws.cell(row = rn+1, column = 3).value='u'
		ws.cell(row = rn+1, column = 4).value='a'
		ws.cell(row = rn+1, column = 5).value='u'
		ws.cell(row = rn+1, column = 6).value='a'
	elif m_wa: #handle wa SC
		ws.cell(row = rn+1, column = 2).value='0' 
		ws.cell(row = rn+1, column = 3).value='u'
		ws.cell(row = rn+1, column = 4).value='a'
		ws.cell(row = rn+1, column = 5).value='u'
		ws.cell(row = rn+1, column = 6).value='a'
	elif m_uo:
		ws.cell(row = rn+1, column = 2).value=m_uo.group(1) 
		ws.cell(row = rn+1, column = 3).value='u'
		ws.cell(row = rn+1, column = 4).value='o'
		ws.cell(row = rn+1, column = 5).value='u'
		ws.cell(row = rn+1, column = 6).value='e'
	elif m_wo: #handle wo SC
		ws.cell(row = rn+1, column = 2).value='0' 
		ws.cell(row = rn+1, column = 3).value='u'
		ws.cell(row = rn+1, column = 4).value='o'
		ws.cell(row = rn+1, column = 5).value='u'
		ws.cell(row = rn+1, column = 6).value='e'
	elif m_uai:
		ws.cell(row = rn+1, column = 2).value=m_uai.group(1) 
		ws.cell(row = rn+1, column = 3).value='u'
		ws.cell(row = rn+1, column = 4).value='ai'
		ws.cell(row = rn+1, column = 5).value='u'
		ws.cell(row = rn+1, column = 6).value='ai'
	elif m_wai: #handle wai SC
		ws.cell(row = rn+1, column = 2).value='0' 
		ws.cell(row = rn+1, column = 3).value='u'
		ws.cell(row = rn+1, column = 4).value='ai'
		ws.cell(row = rn+1, column = 5).value='u'
		ws.cell(row = rn+1, column = 6).value='ai'
	elif m_ui:
		ws.cell(row = rn+1, column = 2).value=m_ui.group(1) 
		ws.cell(row = rn+1, column = 3).value='u'
		ws.cell(row = rn+1, column = 4).value='ei'
		ws.cell(row = rn+1, column = 5).value='u'
		ws.cell(row = rn+1, column = 6).value='ei'
	elif m_wei: #handle wei SC
		ws.cell(row = rn+1, column = 2).value='0' 
		ws.cell(row = rn+1, column = 3).value='u'
		ws.cell(row = rn+1, column = 4).value='ei'
		ws.cell(row = rn+1, column = 5).value='u'
		ws.cell(row = rn+1, column = 6).value='ei'
	elif m_uan:
		ws.cell(row = rn+1, column = 2).value=m_uan.group(1) 
		ws.cell(row = rn+1, column = 3).value='u'
		ws.cell(row = rn+1, column = 4).value='an'
		ws.cell(row = rn+1, column = 5).value='u'
		ws.cell(row = rn+1, column = 6).value='an'
	elif m_wan: #handle wan SC
		ws.cell(row = rn+1, column = 2).value='0'
		ws.cell(row = rn+1, column = 3).value='u'
		ws.cell(row = rn+1, column = 4).value='an'
		ws.cell(row = rn+1, column = 5).value='u'
		ws.cell(row = rn+1, column = 6).value='an'
	elif m_un:
		ws.cell(row = rn+1, column = 2).value=m_un.group(1) 
		ws.cell(row = rn+1, column = 3).value='0'
		ws.cell(row = rn+1, column = 4).value='un'
		ws.cell(row = rn+1, column = 5).value='u'
		ws.cell(row = rn+1, column = 6).value='en'
	elif m_uang:
		ws.cell(row = rn+1, column = 2).value=m_uang.group(1) 
		ws.cell(row = rn+1, column = 3).value='u'
		ws.cell(row = rn+1, column = 4).value='ang'
		ws.cell(row = rn+1, column = 5).value='u'
		ws.cell(row = rn+1, column = 6).value='ang'
	elif m_wang: #handle wang SC
		ws.cell(row = rn+1, column = 2).value='0' 
		ws.cell(row = rn+1, column = 3).value='u'
		ws.cell(row = rn+1, column = 4).value='ang'
		ws.cell(row = rn+1, column = 5).value='u'
		ws.cell(row = rn+1, column = 6).value='ang'
	elif m_ong:
		ws.cell(row = rn+1, column = 2).value=m_ong.group(1) 
		ws.cell(row = rn+1, column = 3).value='0'
		ws.cell(row = rn+1, column = 4).value='ong'
		ws.cell(row = rn+1, column = 5).value='0'
		ws.cell(row = rn+1, column = 6).value='ong'
	elif m_U:
		ws.cell(row = rn+1, column = 2).value=m_U.group(1)
		if m_U.group(1) == 'y': #yu SC
			ws.cell(row = rn+1, column =2).value='0'
		ws.cell(row = rn+1, column = 3).value='0'
		ws.cell(row = rn+1, column = 4).value='v'
		ws.cell(row = rn+1, column = 5).value='0'
		ws.cell(row = rn+1, column = 6).value='v'
	elif m_Ue:
		ws.cell(row = rn+1, column = 2).value=m_Ue.group(1)
		if m_Ue.group(1)=='y': #yue SC
			ws.cell(row = rn+1, column =2).value='0' 
		ws.cell(row = rn+1, column = 3).value='v'
		ws.cell(row = rn+1, column = 4).value='e'
		ws.cell(row = rn+1, column = 5).value='v'
		ws.cell(row = rn+1, column = 6).value='e'
	elif m_Uan:
		ws.cell(row = rn+1, column = 2).value=m_Uan.group(1)
		if m_Uan.group(1) =='y': #yuan SC
			ws.cell(row=rn+1, column=2).value='0' 
		ws.cell(row = rn+1, column = 3).value='v'
		ws.cell(row = rn+1, column = 4).value='an'
		ws.cell(row = rn+1, column = 5).value='v'
		ws.cell(row = rn+1, column = 6).value='an'
	elif m_Un:
		ws.cell(row = rn+1, column = 2).value=m_Un.group(1) 
		if m_Un.group(1)=='y': #yun SC
			ws.cell(row=rn+1, column =2).value='0'
		ws.cell(row = rn+1, column = 3).value='0'
		ws.cell(row = rn+1, column = 4).value='vn'
		ws.cell(row = rn+1, column = 5).value='u'
		ws.cell(row = rn+1, column = 6).value='in'
	elif m_iong:
		ws.cell(row = rn+1, column = 2).value=m_iong.group(1) 
		ws.cell(row = rn+1, column = 3).value='i'
		ws.cell(row = rn+1, column = 4).value='ong'
		ws.cell(row = rn+1, column = 5).value='i'
		ws.cell(row = rn+1, column = 6).value='ong'
	elif m_yong: #yong SC
		ws.cell(row = rn+1, column = 2).value='0' 
		ws.cell(row = rn+1, column = 3).value='i'
		ws.cell(row = rn+1, column = 4).value='ong'
		ws.cell(row = rn+1, column = 5).value='i'
		ws.cell(row = rn+1, column = 6).value='ong'
	elif m_iao:
		ws.cell(row = rn+1, column = 2).value=m_iao.group(1) 
		ws.cell(row = rn+1, column = 3).value='i'
		ws.cell(row = rn+1, column = 4).value='ao'
		ws.cell(row = rn+1, column = 5).value='i'
		ws.cell(row = rn+1, column = 6).value='au'
	elif m_yao: #yao SC
		ws.cell(row = rn+1, column = 2).value='0' 
		ws.cell(row = rn+1, column = 3).value='i'
		ws.cell(row = rn+1, column = 4).value='ao'
		ws.cell(row = rn+1, column = 5).value='i'
		ws.cell(row = rn+1, column = 6).value='au'
	else:
		ws.cell(row = rn+1, column = 2).value='###' 
		ws.cell(row = rn+1, column = 3).value='###'
		ws.cell(row = rn+1, column = 4).value='###'
		ws.cell(row = rn+1, column = 5).value='###'
		ws.cell(row = rn+1, column = 6).value='###'
	wb.save('data.xlsx')