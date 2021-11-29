import urllib2 as url
import sys
#got datas are written with euc-jp.
#localrules are written with shift-jis
localrule=lambda c,n:'http://jbbs.livedoor.jp/'+ str(c) +'/'+str(n)+'/head.txt'
setting=lambda c,n:'http://jbbs.livedoor.jp/bbs/api/setting.cgi/' + str(c) +'/'+str(n)+'/'
subject=lambda c,n:'http://jbbs.livedoor.jp/'+str(c)+'/'+str(n)+'/subject.txt'
getdat=lambda c,n,t:'http://jbbs.livedoor.jp/bbs/rawmode.cgi/'+str(c)+'/'+str(n)+'/'+str(t)+'/'

def dataout(data,enc):
	while True:
		s=data.readline()
		s=unicode(s,enc).encode('utf-8')
		sys.stdout.write(s)
		if len(s)==0:break

if __name__ == '__main__':
	ca='game' #カテゴリ
	bbs='11830' #bbs番号
	tId='1076942817' #スレッド番号
	queryList=(
		(setting(ca,bbs),'euc_jp'),
		(localrule(ca,bbs),'shift_jis'),
		(subject(ca,bbs),'euc_jp'),
		(getdat(ca,bbs,tId),'euc_jp')	)
	for q in queryList:
		print(q)
		try:
			d=url.urlopen(q[0])
		except url.HTTPError,e:
			print e
			continue
		dataout(d,q[1])
