
import pyexcel,os,json,webbrowser,pyrebase
import pyexcel_xlsx
import pyexcel_xls
from win32api import GetSystemMetrics
from tkinter import *
from PIL import ImageTk,Image
from tkinter.filedialog import askopenfilename
from tkinter.filedialog import askopenfilenames
from tkinter import ttk
from tkinter import messagebox
import quickstart
from datetime import datetime
heightofscreen=GetSystemMetrics(1)
widthofscreen=GetSystemMetrics(0)
print(heightofscreen,widthofscreen)
root=Tk()
root.title('Email Sender')
root.iconbitmap('emailsender.ico')
root.geometry('500x480')
Label(text='Email Sender',font=('Helvetica',15,'bold')).place(x=190,y=20)
Label(text='Gmail Authentication',fg='orange',font=('Helvetica',10,'bold underline')).place(x=190,y=60)
html=''
text=''



def heist():
    
    print(ebpath.get())
    print(ipath.get())
    print(xlpath.get())
    print(firebaseconfigpath.get())
    global imagepaths,subject,excelpath,firebaseconfigurationpath,htmlpath,firebase,storage,emailbody1,records
    firebaseconfigurationpath=firebaseconfigpath.get()
    if firebaseconfigurationpath!='':
        if not os.path.exists(firebaseconfigurationpath):
            messagebox.showerror('Error','Invalid path given for firebase configuration')
        else:
            messagebox.showinfo('Firebase','Connecting to Firebase!')
            configkey=open(firebaseconfigurationpath,'r').readlines()
            configkey=configkey[0]
            try:
                firebase=pyrebase.initialize_app(json.loads(configkey))
                messagebox.showinfo('firebase','Connected to firebase successfully!')
                storage=firebase.storage()
            except Exception as e:
                messagebox.showerror('Firebase',e)
                quit()
            
    if xlpath.get()!='':
        if os.path.exists(xlpath.get()):
            excelpath=xlpath.get()
            records=pyexcel.get_records(file_name=excelpath)
        else:
            messagebox.showerror('Error','No such excel file exists')
            quit()
    else:
        messagebox.showwarning('Error','Excel file path is required for getting email recipent data it can\'t be null')
        quit()
    if espath.get()!='':
        subject=espath.get()
        global uploadimagepathorder,uploadimagepathnum
        uploadimagepathnum=[]
        uploadimagepathorder=[]
        if ipath.get()=='':
            if os.path.exists(ebpath.get()):
                htmlpath=ebpath.get()
                emailbody1=emailbodyparser(ebpath.get())
                root.destroy()
                emailpreview(subject,emailbody1)
                
            else:
                messagebox.showerror('Email Body!','No such file exists')
        elif ebpath.get()=='':
           
            imagepaths=(ipath.get()).split(',')
            root.destroy()
            arrangeimage()
    else:
        messagebox.showwarning('Error','Email Subject can\'t be null')

def progressdisplayer(title):
    global root5,progresslblttl,progresslbl,pbar,pcount
    root5=Tk()
    root5.iconbitmap('emailsender.ico')
    root5.title('Email Sender (progress displayer)')
    root5.geometry('400x250')
    Label(root5,text='Email Sender',font=('Helvetica',12,'bold underline')).pack()
    progresslblttl=Label(root5,text=title,font=('Helvetica',10,'bold underline'),fg='orange')
    progresslblttl.pack(pady=20)
    pbar=ttk.Progressbar(root5,orient='horizontal',mode='determinate',length=300)
    pbar.pack(pady=30)
    progresslbl=Label(root5,text='0.0%',fg='green',font=('Helvetica',8,'bold'))
    progresslbl.pack()
    pcount=Label(root5,text='0/0',fg='green',font=('Helvetica',8,'bold'))
    pcount.pack(pady=4)
    root5.update()


def heist1():
    global uploadimagepathorderurl
    count=0
    print('START HEIST PART A')
    print(len(uploadimagepathorder),len(text),len(html))
  
    subject=espath1.get()
    uploadimagepathorderurl=[]
    if len(uploadimagepathorder)==0:
        if text!='':
            
            
            print(subject,emailbody1)
            emailbody=emailbody1
            root3.destroy()
            progressdisplayer('Sending Emails')
   
            for counter in range(len(records)):
                 record=records[count]
                 sendemailbody=emailbody
                
                 for key in varlist:
                     sendemailbody=sendemailbody.replace('#'+key.lower()+'#',str(record[key]))
                     #print(str(record[key]))
                 print(sendemailbody)
                 EmailSender.sendemail(subject,sendemailbody,record['Email'],'text')
                 count+=1    
                 pbar['value']=count/len(records)*100
                 progresslbl.configure(text=str(count/len(records)*100)+'%')
                 pcount.configure(text=str(count)+'/'+str(len(records)))
                 root5.update()
            messagebox.showinfo('Email Sender','Email Sent to all respective recipents!')
            root5.destroy()
            quit()
   
        if html!='':

            emailbody=emailbody1
            root3.destroy()
            progressdisplayer('Sending Emails')
           
            for counter in range(len(records)):
                record=records[count]
                EmailSender.sendemail(subject,emailbody,record["Email"],'html')
                count+=1
                pbar['value']=count/len(records)*100
                progresslbl.configure(text=str(count/len(records)*100)+'%')
                pcount.configure(text=str(count)+'/'+str(len(records)))
                root5.update()
          
            messagebox.showinfo('Email Sender','Email Sent to all respective recipents!')
            root5.destroy()
            quit()
        else:
            print(False)
                
          
 

    elif len(uploadimagepathorder)!=0:
       root3.destroy()
       progressdisplayer('Uploading Images....')
       if firebaseconfigurationpath!='':
         now=datetime.now()
         timestamp=now.strftime('%d-%m-%Y -%H-%M-%S')
         for path in uploadimagepathorder:
             storage.child(timestamp+'/'+os.path.basename(path)).put(path)
             url=storage.child(timestamp+'/'+os.path.basename(path)).get_url(timestamp+'/'+os.path.basename(path))
             uploadimagepathorderurl.append(url)
             count+=1
             pbar['value']=count/len(uploadimagepathorder)*100
             progresslbl.configure(text=str(count/len(uploadimagepathorder)*100)+'%')
             pcount.configure(text=str(count)+'/'+str(len(uploadimagepathorder)))
             
             root5.update()
         print(uploadimagepathorderurl)  
        # messagebox.showinfo('Firebase','All images uploaded and url captured successfully')
         count=0
         root5.destroy()   
         emailbody=imagetohtml(uploadimagepathorderurl)
         print(emailbody)
         progressdisplayer('Sending Email')
      
         for counter in range(len(records)):
            record=records[count]
            EmailSender.sendemail(subject,emailbody,record["Email"],'html')
            count+=1
            pbar['value']=count/len(records)*100
            progresslbl.configure(text=str(count/len(records)*100)+'%')
            pcount.configure(text=str(count)+'/'+str(len(records)))
            root5.update()
          
         messagebox.showinfo('Email Sender','Email Sent to all respective recipents!')
         root5.destroy()
         quit()
       else:
            messagebox.showwarning('Warning','For sending images we\'ll need firebase configuration for storing images on firebase storage')
            quit()
       
    
        


    
  





def imagetohtmltofile():
    global htmlpath
    root2.destroy()
    emailbody=imagetohtml(uploadimagepathorder)

    if os.path.exists('emailbody.html'):
        x=open('emailbody.html','w')
        x.write(emailbody)
        x.close()
        htmlemaibodypreview('file://'+os.path.normpath(os.getcwd())+'/emailbody.html')
    else:
        x=open('emailbody.html','x')
        x.write(emailbody)
        x.close()
        htmlemaibodypreview('file://'+os.path.normpath(os.getcwd())+'/emailbody.html')
    htmlpath='file://'+os.path.normpath(os.getcwd())+'/emailbody.html'
    emailpreview(subject,emailbody)
    
    
    
def htmlemaibodypreview(path):
    try:

        chrome_path=configuration['browserpath']
        webbrowser.register(configuration['browserame'],None,webbrowser.BackgroundBrowser(chrome_path))
        webbrowser.get(configuration['browserame']).open_new_tab(path)
    except Exception as e:
        messagebox.showerror('Browser Error',e)

    

def arrangeimage():
    global libox,image,root2,imageorder
    root2=Tk()
    root2.iconbitmap('emailsender.ico')
    root2.geometry('600x600')
    root2.state('zoomed')
    root2.title('Email Sender (image manager)')
    Label(root2,text='Email Sender',font=('Helvetica',12,'bold underline')).place(x=widthofscreen/2,y=20)
    libox=Listbox(root2,width=400)
    libox.place(x=1,y=60)
    img=Image.open('test.jpg')
    img.resize((100,100))
    photo=ImageTk.PhotoImage(img)
    image=Label(root2,image=photo)
    image.place(x=(widthofscreen/2)-200,y=(heightofscreen/2)-200)
    imageorder=Listbox(root2,width=100)
    imageorder.place(x=widthofscreen/2-300,y=heightofscreen-300)
    imageorder.bind('<<ListboxSelect>>',displayselectedimage1)
    libox.bind('<<ListboxSelect>>',displayselectedimage)
    for k in imagepaths:
        libox.insert(END,os.path.basename(k))
    Button(root2,text="Remove Selected",bg='lime',fg='red',command=removeselected).place(x=widthofscreen/2-165,y=heightofscreen-100)
    Button(root2,text="Clear All",bg='orange',fg='red',command=clearall).place(x=widthofscreen/2-60,y=heightofscreen-100)
    Button(root2,text="Remove 1 before",bg='cyan',fg='red',command=remove1).place(x=widthofscreen/2,y=heightofscreen-100)
    Button(root2,text="Add",bg='yellow',fg='green',command=add).place(x=widthofscreen/2+105,y=heightofscreen-100)
    Button(root2,text=">>>",bg='lime',fg='red',command=imagetohtmltofile,font=('Helvetica',12,'bold underline')).place(x=widthofscreen/2+180,y=heightofscreen-98)
    root2.mainloop()

def emailpreview(subject,body):
    global root3,espath1,ebody,epreviewbtn
    root3=Tk()
    root3.iconbitmap('emailsender.ico')
    root3.title('Email Sender (email preview)')
    root3.geometry('500x500')
    root3.state('zoomed')
    Label(root3,text='Email Sender',font=('Helvetica',12,'bold underline')).place(x=widthofscreen/2,y=20)
    Label(root3,text='Subject:-').place(x=40,y=100)
    espath1=Entry(root3)
    espath1.place(x=105,y=100)
    espath1.insert(0,subject)
    Label(root3,text='Body:-').place(x=40,y=140)
    if html=='':
        ebody=Text(root3,wrap=WORD)
        ebody.place(x=105,y=140)
        ebody.insert(INSERT,body)
        epreviewbtn=Button(root3,text="Get Excel Preview",bg='lime',fg='red',font=('Helvetica',12,'bold underline'),command=exceldatapreview)
        epreviewbtn.place(x=widthofscreen/2+40,y=heightofscreen-98)
        
    else:
        Button(root3,text='Preview email body',bg='lime',fg='red',command=lambda:htmlemaibodypreview(htmlpath),font=('Helvetica',10,'bold')).place(x=105,y=140)
        epreviewbtn=Button(root3,text=">>>>",bg='lime',fg='red',command=heist1,font=('Helvetica',12,'bold underline'))
        epreviewbtn.place(x=widthofscreen/2+40,y=heightofscreen-98)
    root3.mainloop()
def sendpreviewemail():
    if html=='':
        EmailSender.sendemail(subject,ebody.get('1.0','end'),configuration['previewemail'],'text')
        messagebox.showinfo('Email Sender','Preview Email sent')
        epreviewbtn.configure(command=heist1)
        epreviewbtn.configure(text='>>>>')
    


def getchangedbody():
    global emailbody1
    emailbody1=ebody.get(1.0,'end-1c')
def exceldatapreview():
    getchangedbody()
    print(excelpath)
    global record,varlist
    varlist=[]
   
    record=records[0]
    print(record)
   
    body1=ebody.get(1.0,'end-1c')
    for key in record:
        print(key)
        if '#'+key.upper()+'#' in ebody.get(1.0,'end-1c').upper():
            varlist.append(key)
          
            body1=body1.replace('#'+key.lower()+'#',str(record[key]))
    ebody.delete('1.0','end')
    ebody.insert(INSERT,body1)
    print(varlist)
    epreviewbtn.configure(command=sendpreviewemail)
    epreviewbtn.configure(text='send a preview email')
    
    
    








def add():
    try:
        imageorder.insert(END,os.path.basename(imagepaths[libox.curselection()[0]]))
        uploadimagepathorder.append(imagepaths[libox.curselection()[0]])
        uploadimagepathnum.append(libox.curselection()[0])
        print(uploadimagepathnum)
        imageorder.see(len(uploadimagepathorder)-1)
    except Exception as e:
        messagebox.showwarning('Email Sender (image manager)',e)
        
    
def remove1():
    try:
        imageorder.delete(len(uploadimagepathorder)-1,len(uploadimagepathorder))
        uploadimagepathorder.pop(len(uploadimagepathorder)-1)
        uploadimagepathnum.pop(len(uploadimagepathnum)-1)
        imageorder.see(len(uploadimagepathorder)-1)
    except Exception as e:
        messagebox.showwarning('Email Sender (image manager)',e)
def clearall():
    imageorder.delete(0,END)
    uploadimagepathorder.clear()
    uploadimagepathnum.clear()

def removeselected():
    try:
        num=imageorder.curselection()[0]
        num=uploadimagepathnum[num]
        uploadimagepathorder.remove(imagepaths[num])
        uploadimagepathnum.pop(imageorder.curselection()[0])
        imageorder.delete(imageorder.curselection()[0],None)
    except Exception as e:
        messagebox.showwarning('Email Sender (image manager)',e)


def displayselectedimage(event):
    print(imagepaths[libox.curselection()[0]])
    img=Image.open(imagepaths[libox.curselection()[0]])
    img=img.resize((300,300))
    photo=ImageTk.PhotoImage(img)
    image.configure(image=photo)
    image.image=photo
def displayselectedimage1(event):
    img=Image.open(imagepaths[uploadimagepathnum[imageorder.curselection()[0]]])
    img=img.resize((300,300))
    photo=ImageTk.PhotoImage(img)
    image.configure(image=photo)
    image.image=photo

def emailbodyparser(path):
    global html,text
    html=''
    text=''
    x=open(path,'r')
    y=x.readlines()
    if '.txt' in os.path.basename(path):
        for k in y:
            text=text+k
        return(text)
            
    elif '.html' in os.path.basename(path) or '.htm' in os.path.basename(path):
        for k in y:
            html=html+k
        return(html)
def imagetohtml(path):
    global html
    print(path)
    html=''
    for l in path:
        if not l=='':
            html=html+"<p><img src='"+l+"' style='margin-left:auto;margin-right:auto;display:block'/></p>"
    return(html)
    

def emailbodypicker():
    ipathbtn.configure(state='disabled')
    ipath.configure(state='disabled')
    htmlpath=askopenfilename(title='Choose an HTML file/Text file',filetypes=(('HTML','*.html'),('HTML','*.htm'),('Text File','*.txt')))
    ebpath.insert(0,htmlpath)
    htmlpath=htmlpath.replace('/','\\')
    

def imagepicker():
    l=''
    ebpath.configure(state='disabled')
    ebpathbtn.configure(state='disabled')
    imagepaths=askopenfilenames(title='Pick Images',filetypes=(('Images','*.jpg'),('Images','*.jpeg'),('Images','*.png'),('Images','*.gif')))
    imagepaths=list(imagepaths)
    print(imagepaths)
    for k in imagepaths:
        if not l=='':
            l=l+','+k
        else:
            l=k
    ipath.insert(0,l)

def excelfilepicker():
    excelpath=askopenfilename(title='Choose an excel file for email recipeint list',filetypes=(('Excel File','*.xlsx'),))
    xlpath.insert(0,excelpath)

def firebaseconfigurationpicker():
    firebaseconfigp=askopenfilename(title='pick firebase config file',filetypes=(('Javascript Object Notation','*.json'),))
    firebaseconfigpath.insert(0,firebaseconfigp)







if os.path.exists('token.json'):
    try:
        quickstart.main()
    except Exception as e:
        messagebox.showerror('Error',e)
  
    try:
         import EmailSender
    except Exception as e:
        messagebox.showerror('Error',e)


if os.path.exists('application_config.json'):
    configuration=open('application_config.json','r')
    configuration=configuration.readlines()[0]
    configuration=json.loads(configuration)
    print(configuration)



eauthbtn=Button(root,text='Authenticate Access',command=quickstart.main,fg='#D21F3C',font=('Helvetica',8,'bold'))
eauthbtn.place(x=190,y=120)
if os.path.exists('token.json'):
    eauthbtn.destroy()
    eauthbtn=Button(root,text='Authenticate Access',command=quickstart.main,fg='#D21F3C',font=('Helvetica',8,'bold'))
    eauthbtn.place(x=200,y=120)
    eauthbtn.configure(text='Refresh Token',fg='green')
Label(text='Email Config',fg='orange',font=('Helvetica',10,'bold underline')).place(x=190,y=170)
Label(root,text='Email Subject:-').place(x=20,y=200)
espath=Entry(root)
espath.place(x=228,y=200)
Label(root,text='Email Body:-').place(x=20,y=230)
ebpath=Entry(root)
ebpath.place(x=228,y=230)
ebpathbtn=Button(root,text='...',command=emailbodypicker)
ebpathbtn.place(x=365,y=228)
Label(root,text='Image Path:-').place(x=20,y=260)
ipath=Entry(root)
ipath.place(x=228,y=260)
ipathbtn=Button(text='...',command=imagepicker)
ipathbtn.place(x=365,y=258)
Label(root,text='Email Recipeint List Path:-').place(x=20,y=290)
xlpath=Entry(root)
xlpath.place(x=228,y=290)
Button(text='...',command=excelfilepicker).place(x=365,y=288)
Label(text='Firebase Configuration',fg='orange',font=('Helvetica',10,'bold underline')).place(x=190,y=320)
Label(text="Firebase Config File:-").place(x=20,y=360)
firebaseconfigpath=Entry(root)
firebaseconfigpath.place(x=228,y=360)
firebaseconfigpathbtn=Button(root,text='...',command=firebaseconfigurationpicker)
firebaseconfigpathbtn.place(x=365,y=360)
startbtn=Button(root,text='Begin the heist',bg='orange',fg='red',command=heist)
startbtn.place(x=220,y=400)
root.mainloop()
