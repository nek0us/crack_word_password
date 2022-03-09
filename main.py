import sys
import  os
import threading
import queue
import pythoncom
import time
from win32com.client import Dispatch

class docx_crack:
    def __init__(self) -> None:
        self.__messy = 0
        self.__start = time.time()
        self.queue=queue.Queue()
        
    def start_time(self):
        return self.__start
        

    def read_pass(self,passpath):
        f = open(passpath,"r",encoding="utf-8")
        pawd = []
        for x in f.readlines():
            pawd.append(x[:-1])
        print("-------------------------\n  password size:"+"%d len"%(len(pawd)))
        for x in pawd:
            self.queue.put(x)   

    def docc(self,filepath,loc,num):
        while 1:
            thread_title = threading.current_thread().name
            pythoncom.CoInitialize()
            thread_title=Dispatch('Word.Application')
            pythoncom.CoInitialize()
            thread_title.Visible=0
            thread_title.DisplayAlerts=0
            loc.acquire()
            if not self.queue.empty():
                try:
                    
                    pawd = self.queue.get()
                    num1 = num - self.queue.qsize()
                    num2= "%.2f%%" % (num1/num * 100)
                    print("\r  now... "+str(num2), end="")
                    docx = thread_title.Documents.Open(FileName=filepath,ConfirmConversions=False,ReadOnly=True,AddToRecentFiles=False,PasswordDocument=pawd)
                    print("\n  yes,password is "+pawd)
                    self.__messy = 1
                    docx.Close()
                    thread_title.Close()
                    loc.release()
                except:
                    loc.release()
            else:
                loc.release()
                thread_title.Close()
                self.__messy = 1
                break
    def main(self):
        filepath=''
        passpath=''
        threads=[]
        
        loc = threading.Lock()
        try:
            for i in range(len(sys.argv)):
                if ".doc" in sys.argv[i] or ".docx" in sys.argv[i]:
                    if ':\\' in sys.argv[i]:
                        filepath=sys.argv[i]    #获取拖进来的文档路径+名
                    else:
                        filename=sys.argv[i]    #获取文档文件名
                elif ".txt" in sys.argv[i] or ".dict" in sys.argv[i]:
                    if ':\\' in sys.argv[i]:
                        passpath=sys.argv[i]    #获取拖进来的pass路径+名
                    else:
                        passname=sys.argv[i]    #获取pass文件名
            if not filepath:
                filepath=os.getcwd()+"\\"+filename
                passpath=os.getcwd()+"\\"+passname
        except:
            print("\n-------------------------\n未检测到.docx、.doc文档或.txt、.dict密码字典文件\neg：  python3 main.py D:\\..(可选)...\\文档.docx D:\\..(可选)...\\password.txt\n-------------------------\nno .docx, .doc or .txt, .dict detected\neg：  python3 main.py D:\\..(optional)...\\document.docx D:\\..(optional)...\\password.txt\n-------------------------\n")
            exit(0)
        
        self.read_pass(passpath)
        

        print("-------------------------\n  start\n-------------------------")
        num = self.queue.qsize()
        for x in range(4):
            t = threading.Thread(target=self.docc,args=(filepath,loc,num,))
            t.daemon = 1
            threads.append(t)
            
        for t in threads:
            t.start()
            
        while 1:
            if self.__messy == 1:
                break


if __name__ == '__main__':
    start = docx_crack().start_time()
    docx_crack().main()
    print("\n-------------------------\n  end\n-------------------------")
    end = time.time()
    run_time = end - start
    print("  time used: %.5f s" %run_time + "\n-------------------------")