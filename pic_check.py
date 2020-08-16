#识别图片文字
from aip import AipOcr#需要先install baidu-aip才能调用接口
import os,shutil
config={'appId':'xxx','apiKey':'xxxx','secretKey':'xxxx'}
client=AipOcr(**config)
#获取图片信息
def get_pic_content(filepath):
	with open(filepath,'rb') as f:
		return f.read()
#识别图片文字,获取档案信息并以list返回
def img_to_list(img_path):
	img= get_pic_content(img_path)
	res=client.basicGeneral(img)#res为dict
	#print(res)
	result_num=res['words_result_num']
	word_result=res['words_result']#word_result为list，其中每个元素为dict
	#print(word_result)
	dalist=[]
	startnum=8
	while startnum<result_num:#当序号小于list长度时，遍历没个value，查看是否有档案号
		daxx=[]
		try:
			int(word_result[startnum]['words'])
		except ValueError:
			startnum+=1
			continue#跳过此循环执行下次循环
		for i in range(4):
		    daxx.append(word_result[startnum]['words'])
		    startnum+=1
		dalist.extend(daxx)
	#print(dalist)
	return dalist
#遍历文件夹中的信息，并获得档案信息
def read_file_path(path):
	filepaths=[]
	profilelists=[]
	#获得所有文件名
	filenames=[name for name in os.listdir(path) if os.path.isfile(os.path.join(path,name))]
	#获得所有图片的路径
	for name in filenames:
		filepaths.append(os.path.join(path,name))
	#读取图片信息，获得档案信息
	for path in filepaths:
		profilelists.extend(img_to_list(path))
	return filepaths,profilelists
#处理档案
def deal_with_profile(picpath):
	pass
#移动图片
def move_profile(from_path,to_path):
	if not os.path.isfile(from_path):
		print('no such file')
	else:
		tpath,tname=os.path.split(to_path)
		if not os.path.exists(tpath):
			os.mkdir(tpath)
		shutil.move(from_path,to_path)
		#print('move finished')
#处理档案信息主函数
def main():
	path='e:\\pic'#来源图片路径
	tpath='e:\\finishpic'#图片保存路径
	filepaths,profilelists=read_file_path(path)#读取图片路径及信息
	i=0
	while i<len(profilelists):#把每个档案信息拆分出来并做处理
		profile=[]
		for r in range(4):
			profile.append(profilelists[i])
			i+=1
		print(profile)
		#deal_with_profile
	for filepath in filepaths:#移动图片
		path,name=os.path.split(filepath)
		to_path=os.path.join(tpath,name)
		move_profile(filepath,to_path)
		print('move finished')
if __name__=='__main__':
	main()