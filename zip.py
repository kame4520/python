import os
import pathlib
import shutil

tgt_dir = "L:\ほん\wrk\Vuze Downloads\アミグダラ"
os.chdir(tgt_dir)
tgts = os.listdir(path=tgt_dir)

for i in(tgts):
    shutil.make_archive(i,'zip',tgt_dir,i)
    print(tgts)
#shutil.make_archive(tgts,'zip')
print("完了")
