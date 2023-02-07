# -*- coding: utf-8 -*-
"""
Created on Wed Jan 18 16:21:48 2023

@author: Md.Jahidul Islam
"""

try:
    import tkinter as tk                # python 3
    from tkinter import font as tkfont  # python 3
    
except ImportError:
    import Tkinter as tk     # python 2
    import tkFont as tkfont  # python 2
from tkinter import filedialog
import tkinter.messagebox
import warnings
warnings.simplefilter(action='ignore', category=FutureWarning)
import fitz
from PIL import Image
import os
import shutil

class SampleApp(tk.Tk):

    def __init__(self, *args, **kwargs):
        tk.Tk.__init__(self, *args, **kwargs)
        #Main Frame
        self.title("PDF To TIFF OFFLINE CONVERTER")
        MainFrame= tk.Frame(self,bd=10, width=1100, height=400,relief= tk.RIDGE, bg='cadet blue')
        MainFrame.pack(fill=tk.BOTH,expand=1)
        MainFrame.pack_propagate(0)
        
        #Panedwindow Botton and Editor
        BE_pan=tk.PanedWindow(MainFrame,orient=tk.HORIZONTAL)
        
        #Page Frame --Left side of Main Frame
        Page_Frame2= tk.Frame(BE_pan,bd=8, width=800, height=200,relief= tk.RIDGE, bg='AliceBlue')
        Page_Frame2.pack(side="left", fill=tk.BOTH)
        BE_pan.add(Page_Frame2)
        
        #Main ELEMENT Frame --Right side of Main Frame
        Main_ELEMENT_Frame5= tk.Frame(BE_pan,bd=8, width=50, height=200,relief= tk.RIDGE, bg='AliceBlue')
        Main_ELEMENT_Frame5.pack(side="right")
        Main_ELEMENT_Frame5.pack_propagate(0)
        
        BE_pan.add(Main_ELEMENT_Frame5)
        
        BE_pan.pack(fill = tk.BOTH, expand = True)
          
        # This method is used to show sash
        BE_pan.configure(sashrelief = tk.RAISED)
        
        #Panedwindow Element and display
        ED_pan=tk.PanedWindow(Page_Frame2,orient=tk.VERTICAL)
        #Page Frame --top side of Main Frame
        Page_Frame1= tk.Frame(ED_pan,bd=8, width=800, height=100,relief= tk.RIDGE, bg='cadet blue')
        Page_Frame1.pack(side="top", fill=tk.BOTH)
        ED_pan.add(Page_Frame1)
        
        #Main ELEMENT Frame --bottom side of Main Frame
        Frame5= tk.Frame(ED_pan,bd=8, width=800, height=100,relief= tk.RIDGE, bg='cadet blue')
        Frame5.pack(side="bottom")
        Frame5.pack_propagate(0)
        
        ED_pan.add(Frame5)
        ED_pan.pack(fill = tk.BOTH, expand = True)
        ED_pan.configure(sashrelief = tk.RAISED)
        
        
        Fact = """\nHere We Can Get The Update Of The Runing Process."""
        #=============function============
        def Open():
            global path
            path= filedialog.askopenfilenames(title="Select PDF FIles",filetypes=(("PDF Files","*.pdf"),("Excel Files","*.xlsx"),("CSV Files","*.csv"),("All Files","*.*")))
            nb_file="\nNumber of files selected: "+str(len(path))
            T.insert(tk.END,nb_file)
            for i in path:
                filename = i.split('/')[len(i.split('/'))-1]
                filename =filename.replace(".pdf","")
                filename1="\n  "+filename
                T.insert(tk.END,filename1)
                
            
            
        #Location to Save the the report
        def saving():
            global save_only
            save_only= filedialog.askdirectory(title="Select the output folder")
            T.insert(tk.END,f"\nSelected Folder To Save The TIFF File:{save_only}")
        
        def slider_changed(event):
            pass
            
            
        def start():
            for pdf1 in path:
                #Creat Forlder
                dpi = DPI.get()  
                zoom = dpi / Zoom.get()
                magnify = fitz.Matrix(zoom, zoom)  
                doc = fitz.open(pdf1)
                filename = pdf1.split('/')[len(pdf1.split('/'))-1]
                filename =filename.replace(".pdf","")
                filename1="\n    Start Converting For:"+filename
                T.insert(tk.END,filename1)
                
                
                # path
                path_mak = os.path.join(save_only, filename)
                try: 
                    os.mkdir(path_mak) 
                except OSError as error:
                    tkinter.messagebox.showerror('Error', 'Error: Please Select The location To Save The Files!')
                     
                
                add=[]
                for index, page in enumerate(doc):
                    pix = page.get_pixmap(matrix=magnify)  
                    pix.save(f"{path_mak}/page-{page.number}.tiff")
                    im1 = Image.open(f"{path_mak}/page-{page.number}.tiff")
                    if index==0:
                        pass
                    else:
                        add.append(im1)
                        
                    
                    
                im1 = Image.open(f"{path_mak}/page-0.tiff")
                output_tiff=f"{save_only}/{filename}.tiff"
                im1.save(output_tiff, save_all=True, append_images=add, compression='tiff_lzw', tiffinfo={317: 2})
                T.insert(tk.END,f"\n{filename} Is Converted")
                # removing directory
                shutil.rmtree(path_mak, ignore_errors=False)
            os.startfile(save_only)
            tkinter.messagebox.showinfo("PDF TO TIFF","ALL The Pdfs Are Converted Successfully")
                
                        
                
        #===========Button=========
        pdf_button= tk.Button(Main_ELEMENT_Frame5,font =('Bahnschrift Light SemiCondensed', 12, 'bold'),text="SELECT PDFs", bd =5, command= Open).pack(fill=tk.BOTH, expand=1)
        Save_button= tk.Button(Main_ELEMENT_Frame5,font =('Bahnschrift Light SemiCondensed', 12, 'bold'),text="SELECT THE LOCATION TO SAVE", bd =5, command= saving).pack(fill=tk.BOTH, expand=1)
        Convert_button= tk.Button(Main_ELEMENT_Frame5,font =('Bahnschrift Light SemiCondensed', 12, 'bold'),text="Start Converting", cursor="plus",activeforeground='#58D68D',activebackground='#5D6D7E', bd =5, command= start).pack(fill=tk.BOTH, expand=1)
        #===========text=======
        # Create text widget and specify size.
        T = tk.Text(Frame5)
        T.insert(tk.END, Fact)
        T.pack(fill=tk.BOTH, expand=1)
        #================ scorol bar========
        DPI= tk.Scale(Page_Frame1,from_= 300, to=600,orient='horizontal',command=slider_changed)
        DPI.pack(fill=tk.BOTH, expand=1)
        DPIlable=tk.Label(Page_Frame1, text= "Please Select DPI  Value\n Default value: 300", bd =2)
        DPIlable.pack(fill=tk.BOTH, expand=1)
        Zoom= tk.Scale(Page_Frame1,from_= 72, to=200,orient='horizontal',command=slider_changed)
        Zoom.pack(fill=tk.BOTH, expand=1)
        Zoomlable=tk.Label(Page_Frame1, text= "Please Select Zoom Factor Value\nDefault value: 72\nSelect 120-180 for best output", bd =2)
        Zoomlable.pack(fill=tk.BOTH, expand=1)

if __name__ == "__main__":
    app = SampleApp()
    app.mainloop()