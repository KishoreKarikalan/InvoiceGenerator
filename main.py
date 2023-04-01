""""
Company Address must contain 2 lines.
Customer Address must contain 3 lines.
"""

import json
import tkinter as tk
import pandas as pd
import tkinter.ttk as ttk
import datetime as dt
from docxtpl import DocxTemplate
from tkinter import messagebox

invoice_products = {}
invoice_buyer : str
invoice_number : int
interstate_fl : bool

class WelcomeScreen():

    def __init__(self,window : tk.Tk):
        self.window = window
        window.title("Make Invoice")
        window.geometry("1920x1080+0+0")

    def clear_screen(func):
        def wrapper_function(self):
            frames = self.window.winfo_children()
            for frame in frames:
                if isinstance(frame, tk.Frame):
                    frame.destroy()
            func(self)
        return wrapper_function

    @clear_screen
    def load_welcome(self):
        #Frame1
        frame1 = tk.Frame(self.window, bg='black', height=600, width=1080)
        frame1.pack(side='top', fill='both')
        frame1.pack_propagate(False)  # Does not shrink frame to fit content.
        content = tk.Label(frame1, text="Build,Print and Export Invoices.", font=('Arial', 32, ''), fg='white',
                           bg='black')
        content.pack(side='bottom')
        content.config(pady=80)
        welcome = tk.Label(frame1, text="WELCOME", font=('Arial', 65, 'bold'), fg='white', bg='black')
        welcome.pack(side='bottom')
        # Frame2
        frame2 = tk.Frame(self.window, bg='white', height=1360, width=1080,padx=290,pady=180)
        frame2.pack(side='top', fill='both')
        frame2.grid_propagate(False)
        address_book = tk.Button(frame2, text="Address Book", command=self.load_customer,height = 2,width=20,font=('Arial', 16, 'bold'),bg = 'black',fg = 'white')
        product_list = tk.Button(frame2, text="Product List", command=self.load_product,height = 2,width=20,font=('Arial', 16, 'bold'),bg = 'black',fg = 'white')
        generate_invoice = tk.Button(frame2, text="Generate Invoice", command=self.load_generate,height = 2,width=20,font=('Arial', 16, 'bold'),bg = 'black',fg = 'white')
        address_book.grid(row = 0,column=0,padx=80)
        product_list.grid(row = 0,column=1,padx=80)
        generate_invoice.grid(row=0, column=80,padx=80)
        #profile button
        profile_button = tk.Button(frame1,text="Profile",command=self.load_company_profile, height=1, width=10, bg='white', fg='black',
                           font=('Arial', 12, 'bold'))
        profile_button.place(x=100,y=40)

        exit_button = tk.Button(frame1,text="Exit",command=exit_app, height=1, width=10, bg='white', fg='black',
                           font=('Arial', 12, 'bold'))
        exit_button.place(x=1700,y=40)

    @clear_screen
    def load_customer(self) :
        # Frame 1
        frame1 = tk.Frame(self.window, bg='black', height=600, width=1080)
        frame1.pack(side='top', fill='both')
        frame1.pack_propagate(False)  # Does not shrink frame to fit content.
        customer_detail_label = tk.Label(frame1, text="Customer Details", fg='white', bg='black',
                                        font=('Arial', 30, 'bold'))
        customer_detail_label.place(x=80, y=130)
        # Frame 2
        frame2 = tk.Frame(self.window, bg='white', height=1360, width=1080, padx=270, pady=80)
        frame2.pack(side='top', fill='both')
        frame2.grid_propagate(False)

        # Table storing product data.
        s = ttk.Style()
        s.theme_use('clam')

        # Add the rowheight
        s.configure('Treeview', rowheight=100, font=("Arial", 14))
        s.configure('Treeview.Heading', font=("Arial", 17))
        tree = ttk.Treeview(self.window, columns=("col1","col2","col3"), height=7)
        tree.place(x=150, y=200)
        tree.heading("#0", text="Buyer Name")
        tree.heading("col1", text="Address")
        tree.heading("col2", text="GST No")
        tree.heading("col3", text="phone")
        i = 0
        for customer in customer_data:
            tree.insert("", i, text=customer, values=(customer_data[customer][0],customer_data[customer][1],customer_data[customer][2]))
            i += 1

        def add_item():
            customer_data[customer_name.get()] = [address.get("1.0","end-1c"),gst.get(),phone.get()]
            self.load_customer()

        def delete_item():
            del customer_data[customer_name.get()]
            self.load_customer()

        # Add Form Box
        form_box = tk.Frame(self.window, height=400, width=900, pady=40, padx=10, bg='white')
        form_box.place(x=1000, y=400)
        form_box.grid_propagate(False)

        customer_name_label = tk.Label(form_box, text="Customer Name", bg='white', font=("Arial", 16, ""))
        customer_name_label.grid(row=0, column=0, padx=10, pady=10)
        customer_name = tk.Entry(form_box, font=("Arial", 14, ""))
        customer_name.grid(row=1, column=1)

        gst_label = tk.Label(form_box, text="GST Number", bg='white', font=("Arial", 16, ""))
        gst_label.grid(row=0, column=2, padx=10, pady=10)
        gst = tk.Entry(form_box, font=("Arial", 14, ""))
        gst.grid(row=1, column=3)

        phone_label = tk.Label(form_box, text="Phone Number", bg='white', font=("Arial", 16, ""))
        phone_label.grid(row=2, column=0, padx=10, pady=10)
        phone = tk.Entry(form_box, font=("Arial", 14, ""))
        phone.grid(row=3, column=1)

        address_label = tk.Label(form_box, text="Buyer Address", bg='white', font=("Arial", 16, ""))
        address_label.grid(row=2, column=2, padx=10, pady=10)
        address = tk.Text(form_box, height=8, width=24)
        address.grid(row=3, column=3,rowspan=2)
        # Add Button.
        add = tk.Button(form_box, text="Add Customer", bg='black', fg='white', font=("Arial", 16, ""), width=16,
                        command=add_item)
        add.grid(row=4,column=0,pady=20,padx=10)
        # Delete Button.
        delete = tk.Button(form_box, text="Remove Customer", bg='black', fg='white', font=("Arial", 16, ""), width=16,
                           command=delete_item)
        delete.grid(row=4,column=1,pady=20,padx=10)

        def save():
            self.load_welcome()

        # Submit
        submit = tk.Button(frame2, text="Save & Exit", command=save, height=3, width=20, bg='black', fg='white',
                           font=('Arial', 12, 'bold'))
        submit.place(x=1300, y=220)

    @clear_screen
    def load_company_profile(self):
        # Frame 1
        frame1 = tk.Frame(self.window, bg='black', height=600, width=1080)
        frame1.pack(side='top', fill='both')
        frame1.pack_propagate(False)  # Does not shrink frame to fit content.
        company_info_label = tk.Label(frame1,text = "Company Details",fg='white',bg='black',font=('Arial', 18, 'bold'))
        company_info_label.place(x=250,y=60)
        # Frame 2
        frame2 = tk.Frame(self.window, bg='white', height=1360, width=1080, padx=290, pady=180)
        frame2.pack(side='top', fill='both')
        frame2.grid_propagate(False)
        bank_info_label = tk.Label(self.window, text="Bank Details", fg='black', bg='white',
                                      font=('Arial', 18, 'bold'))
        bank_info_label.place(x=250, y=650)
        # Form Box
        form_box = tk.Frame(self.window, height = 400,width=800,pady=50,padx=60,bg='white')
        form_box.place(x=250,y=100)
        form_box.grid_propagate(False)

        text_font1 = ('Arial', 12, '')
        text_font2 = ('Arial', 10, '')

        company_name_label = tk.Label(form_box,text='Company name',font=text_font1,bg='white')
        company_name_label.grid(row=0,column=0,padx=10,pady=10)
        company_name = tk.Entry(form_box,font=text_font2)
        company_name.insert(0,company_profile_dict["Company Name"][0])
        company_name.grid(row=1,column=1,padx=10,pady=10)

        gst_label = tk.Label(form_box, text='GST Number',font=text_font1,bg='white')
        gst_label.grid(row=0, column=2,padx=10,pady=10)
        gst_no = tk.Entry(form_box,font=text_font2)
        gst_no.insert(0, company_profile_dict["GST No"][0])
        gst_no.grid(row=1, column=3,padx=10,pady=10)

        telephone_label = tk.Label(form_box, text='Telephone Number',font=text_font1,bg='white')
        telephone_label.grid(row=2, column=0,padx=10,pady=10)
        telephone = tk.Entry(form_box,font=text_font2)
        telephone.insert(0, company_profile_dict["Telephone No"][0])
        telephone.grid(row=3, column=1,padx=10,pady=10)

        address_label = tk.Label(form_box, text='Company Adress',font=text_font1,bg='white')
        address_label.grid(row=2, column=2, padx=10, pady=10)
        address = tk.Text(form_box, font=text_font2,height=8,width=24)
        address.insert("1.0",company_profile_dict["address"][0])
        address.grid(row=3, column=3, padx=10, pady=10,rowspan=3)

        #Bank Details
        form_box2 = tk.Frame(self.window, height=300, width=800, pady=60, padx=60)
        form_box2.place(x=250, y=700)
        form_box2.grid_propagate(False)

        bank_label = tk.Label(form_box2, text='Bank Name',font=text_font1)
        bank_label.grid(row=0, column=0,padx=10,pady=10)
        bank_name = tk.Entry(form_box2,font=text_font2)
        bank_name.insert(0, company_profile_dict["Bank"][0])
        bank_name.grid(row=1, column=1,padx=10,pady=10)

        branch_label = tk.Label(form_box2, text='Bank Branch',font=text_font1)
        branch_label.grid(row=0, column=2,padx=100,pady=10)
        bank_branch = tk.Entry(form_box2,font=text_font2)
        bank_branch.insert(0, company_profile_dict["Branch"][0])
        bank_branch.grid(row=1, column=3,padx=10,pady=10)

        account_number_label = tk.Label(form_box2, text='Account Number',font=text_font1)
        account_number_label.grid(row=2, column=0,padx=10,pady=10)
        acc_no = tk.Entry(form_box2,font=text_font2)
        acc_no.insert(0, company_profile_dict["Account No"][0])
        acc_no.grid(row=3, column=1,padx=10,pady=10)

        ifsc_label = tk.Label(form_box2, text='IFSC Code ',font=text_font1)
        ifsc_label.grid(row=2, column=2,padx=10,pady=10)
        ifsc_code = tk.Entry(form_box2,font=text_font2)
        ifsc_code.insert(0, company_profile_dict["IFSC code"][0])
        ifsc_code.grid(row=3, column=3,padx=10,pady=10)

        igst_label = tk.Label(form_box, text='IGST percentage',font=text_font1,bg='white')
        igst_label.grid(row=4, column=0,padx=10,pady=10)
        igst_val = tk.Entry(form_box,font=text_font2)
        igst_val.insert(0, company_profile_dict["IGST val"][0])
        igst_val.grid(row=5, column=1,padx=10,pady=10)

        def save():
            if (not (str(igst_val.get()).isnumeric())):
                if(int(igst_val.get()) not in range(0,101)):
                    messagebox.showerror("IGST Percentage","IGST Percentage must be a number between 1 to 100")
                    return
            company_profile_dict['Company Name'][0] = company_name.get()
            company_profile_dict['GST No'][0] = gst_no.get()
            company_profile_dict['Telephone No'][0] = telephone.get()
            company_profile_dict['Bank'][0] = bank_name.get()
            company_profile_dict['Branch'][0] = bank_branch.get()
            company_profile_dict['Account No'][0] = acc_no.get()
            company_profile_dict['IFSC code'][0] = ifsc_code.get()
            company_profile_dict['IGST val'][0] = igst_val.get()
            company_profile_dict['address'][0] = address.get("1.0","end-1c")
            root.load_welcome()

        # Submit
        submit = tk.Button(frame2, text="Save & Exit", command=save,height=3,width=20,bg = 'black',fg = 'white',font=('Arial', 12, 'bold'))
        submit.place(x=1200, y=100)

    @clear_screen
    def load_product(self) :
        # Frame 1
        frame1 = tk.Frame(self.window, bg='black', height=600, width=1080)
        frame1.pack(side='top', fill='both')
        frame1.pack_propagate(False)  # Does not shrink frame to fit content.
        product_detail_label = tk.Label(frame1, text="Product Details", fg='white', bg='black',
                                      font=('Arial', 25, 'bold'))
        product_detail_label.place(x=250, y=250)
        # Frame 2
        frame2 = tk.Frame(self.window, bg='white', height=1360, width=1080, padx=270, pady=80)
        frame2.pack(side='top', fill='both')
        frame2.grid_propagate(False)

        #Table storing product data.
        s = ttk.Style()
        s.theme_use('clam')

        # Add the rowheight
        s.configure('Treeview', rowheight=60,font=("Arial",14,"bold"),backgroud='white')
        s.configure('Treeview.Heading', font=("Arial",17))

        tree = ttk.Treeview(self.window,columns=("col1"),height=10)
        tree.place(x=350, y=300)
        tree.heading("#0", text="Product Name")
        tree.heading("col1", text="Price")
        i=0
        for item_name in product_data:
             tree.insert("",i,text=item_name,values=(product_data[item_name]))
             i+=1

        def add_item():
            if not (str(product_price.get()).isnumeric()):
                messagebox.showerror("Product Price","Product Price must be a number.")
                return
            product_data[product_name.get()] = product_price.get()
            self.load_product()

        def delete_item():
            del product_data[product_name.get()]
            self.load_product()

        #Add Form Box
        form_box = tk.Frame(self.window, height=200, width=790, pady=50, padx=20, bg='white')
        form_box.place(x=1000, y=400)
        form_box.grid_propagate(False)

        product_name_label = tk.Label(form_box,text="Product Name",bg='white',font=("Arial",16,""))
        product_name_label.grid(row = 0, column=0,padx=10,pady=10)
        product_name = tk.Entry(form_box,font=("Arial",14,""))
        product_name.grid(row = 1, column=1)

        product_price_label = tk.Label(form_box, text="Product Price",bg='white',font=("Arial",16,""))
        product_price_label.grid(row = 0, column=2,padx=10,pady=10)
        product_price = tk.Entry(form_box,font=("Arial",14,""))
        product_price.grid(row = 1, column=3)
        # Add Button.
        add = tk.Button(frame2,text="Add Product",bg='black',fg='white',font=("Arial",16,""),width=16,command=add_item)
        add.place(x=850,y=0)
        # Delete Button.
        delete = tk.Button(frame2, text="Remove Product",bg='black',fg='white',font=("Arial",16,""),width=16,command=delete_item)
        delete.place(x=1200,y=0)

        def save():
           self.load_welcome()
        # Submit
        submit = tk.Button(frame2, text="Save & Exit", command=save,height=3,width=20,bg = 'black',fg = 'white',font=('Arial', 12, 'bold'))
        submit.place(x=1300, y=220)

    @clear_screen
    def load_generate(self):

        product_lst = list(product_data.keys())
        customer_lst = list(customer_data.keys())

        def load_table():
            pass
        def search_product(event):
            value = event.widget.get()  #To get Content written in box
            if value == '': #if no character is typed display all
                product['values'] = product_lst
            else:
                data = []
                for item in product_lst:
                    if item.lower() == value.lower():
                        data.append(item)
                product['values'] = data
        def search_buyer(event):
            value = event.widget.get()
            if value == '':
                buyer['values'] = customer_lst
            else:
                data = []
                for item in customer_lst:
                    if item.lower() == value.lower():
                        data.append(item)
                buyer['values'] = data

        def load_price():
            price.delete(0, tk.END)
            price.insert(0,product_data[product.get()])

        # Frame 1
        frame1 = tk.Frame(self.window, bg='black', height=600, width=1080)
        frame1.pack(side='top', fill='both')
        frame1.pack_propagate(False)  # Does not shrink frame to fit content.
        product_detail_label = tk.Label(frame1, text="Invoice Products", fg='white', bg='black',
                                        font=('Arial', 25, 'bold'))
        product_detail_label.place(x=150, y=100)
        # Frame 2
        frame2 = tk.Frame(self.window, bg='white', height=1360, width=1080, padx=270, pady=80)
        frame2.pack(side='top', fill='both')
        frame2.grid_propagate(False)

        #Form
        form = tk.Frame(self.window,bg='white',height=500,width=780,padx=40,pady=80)
        form.place(x=1050,y=150)
        form.grid_propagate(False)

        #search box for Buyer
        buyer_label = tk.Label(form,text="Buyer Name",font=("Arial",16,""),bg='white')
        buyer_label.grid(row=0,column=0,padx=5,pady=15)
        buyer = ttk.Combobox(form,values=list(customer_data.keys()),font=("Arial",12,""))
        buyer.set("Buyer")
        buyer.bind("<KeyRelease>",search_buyer)
        buyer.grid(row=1,column=1,padx=5,pady=15)

        #invoice number
        inv_no_label = tk.Label(form, text="Invoice Number",bg='white',font=("Arial",16,""))
        inv_no_label.grid(row=0, column=2,padx=5,pady=15)
        inv_no = tk.Entry(form,font=("Arial",12,""))
        inv_no.grid(row=1, column=3,padx=5,pady=15)

        # search box for product
        product_label = tk.Label(form, text="Product Name", font=("Arial", 16, ""), bg='white')
        product_label.grid(row=2, column=0,padx=5,pady=15)
        product = ttk.Combobox(form, values=list(product_data.keys()), font=("Arial", 12, ""))
        product.set("Product")
        product.bind("<KeyRelease>", search_product)
        product.grid(row=3, column=1,padx=5,pady=15)

        # product price
        price_label = tk.Label(form, text="Product Price", bg='white', font=("Arial", 16, ""))
        price_label.grid(row=2, column=2,padx=5,pady=15)
        price = tk.Entry(form, font=("Arial", 12, ""))
        price.grid(row=3, column=3,padx=5,pady=15)

        # product price
        qty_label = tk.Label(form, text="Product Quantity", bg='white', font=("Arial", 16, ""))
        qty_label.grid(row=4, column=0, padx=5, pady=15)
        qty = tk.Entry(form, font=("Arial", 12, ""))
        qty.insert(0,"1")
        qty.grid(row=5, column=1, padx=5, pady=15)

        #load price
        load = tk.Button(form,text="load price",font=("Arial", 14, ""),bg='black',fg='white',padx=20,command=load_price)
        load.grid(row=5,column=3,padx=10)

        #interstate checkbox
        var = tk.IntVar()
        checkbox = tk.Checkbutton(frame2, text="Transaction Outside State.", variable=var,font=("Arial", 14, ""),bg='white')
        checkbox.place(x=1000,y=100)

        #Submit
        def save():
            global interstate_fl,invoice_buyer,invoice_number
            interstate_fl = var.get()
            invoice_buyer = buyer.get()
            invoice_number = inv_no.get()
            generate_invoice()
            self.load_welcome()
        # Submit
        submit = tk.Button(frame2, text="Generate", command=save,height=3,width=20,bg = 'black',fg = 'white',font=('Arial', 12, 'bold'))
        submit.place(x=1300, y=220)

        def add_item():
            if not (str(price.get()).isnumeric()):
                messagebox.showerror("Product Price","Product Price must be a number.")
                return
            if not (str(qty.get()).isnumeric()):
                messagebox.showerror("Product Quantity","Product Quantity must be a number.")
                return
            invoice_products[product.get()] = [price.get(),qty.get()]
            load_table()
        def delete_item():
            del invoice_products[product.get()]
            load_table()
        def load_table():
            # Table storing product data.
            s = ttk.Style()
            s.theme_use('clam')

            # Add the rowheight
            s.configure('Treeview', rowheight=60, font=("Arial", 14, "bold"), backgroud='white')
            s.configure('Treeview.Heading', font=("Arial", 17))

            tree = ttk.Treeview(self.window, columns=("col1","col2"), height=10)
            tree.place(x=240, y=160)
            tree.heading("#0", text="Product Name")
            tree.heading("col1", text="Price")
            tree.heading("col2", text="Qty")
            i = 0
            for item_name in invoice_products:
                tree.insert("", i, text=item_name, values=(invoice_products[item_name][0],invoice_products[item_name][1]))
                i += 1

        # Add Button.
        add = tk.Button(frame2, text="Add Product", bg='black', fg='white', font=("Arial", 16, ""), width=16,
                        command=add_item)
        add.place(x=850, y=0)
        # Delete Button.
        delete = tk.Button(frame2, text="Remove Product", bg='black', fg='white', font=("Arial", 16, ""),
                           width=16, command=delete_item)
        delete.place(x=1200, y=0)


def number_to_word(number):
    if number == 0:
        return "zero"
    if number < 0:
        return "minus " + number_to_word(abs(number))
    if number < 20:
        return ["One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine", "Ten", "Eleven", "Twelve", "Thirteen", "Fourteen", "Fifteen", "Sixteen", "Seventeen", "Eighteen", "Nineteen"][number-1]
    if number < 100:
        return [None, None, "twenty", "thirty", "forty", "fifty", "sixty", "seventy", "eighty", "ninety"][number//10] + ('' if number%10 == 0 else ' ' + number_to_word(number%10))
    if number < 1000:
        return number_to_word(number//100) + " hundred" + (' and' if number%100 != 0 else '') +('' if number%100 == 0 else ' ' + number_to_word(number%100))
    if number < 100000:
        return number_to_word(number//1000) + " thousand" + ('' if number%1000 == 0 else ' ' + number_to_word(number%1000))
    if number < 10000000:
        return number_to_word(number//100000) + " Lakh" + ('' if number%100000 == 0 else ' ' + number_to_word(number%100000))

def exit_app():

    with open('products.json','w') as f:
        json.dump(product_data,f)
    with open('customers.json','w') as f:
        json.dump(customer_data,f)
    df = pd.DataFrame(company_profile_dict, index=[0])
    df.to_json('profile.json', orient='records')
    exit()

def generate_invoice() :
    #global invoice_buyer

    doc = DocxTemplate('InvoiceTemplate.docx')

    invoice_dict = {}

    invoice_dict['company_name'] = company_profile_dict['Company Name'][0]
    invoice_dict['address1'] = (company_profile_dict['address'][0]).split('\n')[0]
    print((company_profile_dict['address'][0]).split('\n')[0])
    invoice_dict['address2'] = (company_profile_dict['address'][0]).split('\n')[1]
    invoice_dict['phone'] = company_profile_dict['Telephone No'][0]
    invoice_dict['gst_no'] = company_profile_dict['GST No'][0]

    invoice_dict['address'] = '\n\t\t'.join(customer_data[invoice_buyer][0].split('\n')) + '\n\t\t'
    invoice_dict['address'] += 'phone :'+str(customer_data[invoice_buyer][1]) + '\n\t\t'
    invoice_dict['address'] += 'GST no :'+str(customer_data[invoice_buyer][2])
    invoice_date = dt.datetime.now()
    invoice_dict['Date'] = invoice_date.strftime("%d-%m-%y") #To format the date as string
    invoice_dict['inv_no'] = invoice_number
    products = []
    i,total_val,total_qty=1,0,0
    for product in invoice_products:
        qty = int(invoice_products[product][1])
        price = int(invoice_products[product][0])
        total_per_item = qty*price
        products.append([i,product,84137010,qty,price,total_per_item])
        total_val += total_per_item
        total_qty += qty
        i+=1

    invoice_dict['invoice_list'] = products
    igst_perc,sgst_perc,cgst_perc,igst_val,sgst_val,cgst_val = 0,0,0,0,0,0

    if interstate_fl :
        igst_perc = float(company_profile_dict['IGST val'][0])
        igst_val = total_val*(igst_perc/100)
    else:
        sgst_perc = int(company_profile_dict['IGST val'][0])/2
        sgst_val = total_val * (sgst_perc / 100)
        cgst_perc = float(company_profile_dict['IGST val'][0])/2
        cgst_val = total_val * (cgst_perc / 100)

    invoice_dict['sub_t'] = format(total_val,'.2f')

    invoice_dict['sgst_p'] = str(sgst_perc) + '%'
    invoice_dict['sgst_v'] = format(sgst_val,'.2f')
    invoice_dict['cgst_p'] = str(cgst_perc) + '%'
    invoice_dict['cgst_v'] = format(cgst_val,'.2f')
    invoice_dict['igst_p'] = str(igst_perc) + '%'
    invoice_dict['igst_v'] = format(igst_val,'.2f')

    invoice_dict['amt'] = format(round((total_val + igst_val + sgst_val + cgst_val),0),'.2f')
    # invoice_dict['qty'] = total_qty

    invoice_dict['amt_word'] = number_to_word(int(round((total_val + igst_val + sgst_val + cgst_val),0))) + ' oniy'

    invoice_dict['bank'] = company_profile_dict[ "Bank"][0]
    invoice_dict['branch'] = company_profile_dict["Branch"][0]
    invoice_dict['acc_no'] = company_profile_dict["Account No"][0]
    invoice_dict['ifsccode'] = company_profile_dict["IFSC code"][0]

    doc.render(invoice_dict)
    doc.save('new_invoice.docx')


if __name__ == '__main__' :

    window = tk.Tk()
    root = WelcomeScreen(window)

    try:
        company_profile_dict = (pd.read_json('profile.json')).to_dict()

        with open('products.json','r') as f :
            product_data = json.load(f)
        with open('customers.json','r') as f :
            customer_data = json.load(f)
        root.load_welcome()

    except FileNotFoundError :
        company_profile_dict_keys = ["Company Name", "GST No", "Telephone No", "Bank", "Branch", "Account No",
                                     "IFSC code", "IGST val", "address"]
        company_profile_dict = {key: f"Enter {key}" for key in company_profile_dict_keys}
        product_data = {}
        customer_data = {}
        root.load_company_profile()

    window.mainloop()





