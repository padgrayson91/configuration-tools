import os
import os.path
import sys
import xlrd
import csv
import json

if sys.version_info[0] < 3:
	from Tkinter import _setit
	from Tkinter import *
	from ttk import *
	from tkFileDialog import askdirectory, askopenfilename
	from tkMessageBox import showwarning, showinfo
else:
	from tkinter import _setit
	from tkinter import *
	from tkinter.ttk import *
	from tkinter.filedialog import askdirectory, askopenfilename
	from tkinter.messagebox import showwarning, showinfo

class ExcelConverterDialog(Frame):
	def __init__(self, parent):
		Frame.__init__(self, parent)

		self.parent = parent
		self.initUI()

	def add_mapping(self):
		self.conversions.append({"source":StringVar(), "destination":StringVar(), "delimeter":StringVar()})
		conversion_frame = Frame(self.conversions_frame)
		self.conversion_frames.append(conversion_frame)
		conversion_box = apply(OptionMenu, (conversion_frame, self.conversions[len(self.conversions) - 1]["source"]) + tuple(self.input_column_headers))
		self.conversion_boxes.append(conversion_box)
		conversion_box.pack(side=LEFT)
		output_box = apply(OptionMenu, (conversion_frame, self.conversions[len(self.conversions) - 1]["destination"]) + tuple(self.output_column_headers))
		self.output_boxes.append(output_box)
		label = Label(conversion_frame, text="  map column to  ")
		label.pack(side=LEFT)
		output_box.pack(side = LEFT)
		delim_label = Label(conversion_frame, text= "  Delimeter (for combined entries):")
		delim_label.pack(side=LEFT)
		delimeter_entry = Entry(conversion_frame, textvar=self.conversions[len(self.conversions) - 1]["delimeter"])
		self.delimeter_entries.append(delimeter_entry)
		delimeter_entry.pack(side=LEFT)
		conversion_frame.pack(side=TOP, fill=X)

	def do_load(self, conversion_name):
		f = open("mappings.json", "r")
		lines = "".join(f.readlines())
		f.close()

		configs_json = json.loads(lines)
		conversion = configs_json[conversion_name]

		self.load_mapping(conversion)

	def load_mapping(self, json_in):
		for i in range(len(self.conversion_frames)):
			self.remove_mapping()

		conversions = json_in["conversions"]
		self.constants = json_in["constants"]


		for item in conversions:
			self.add_mapping()

		for conversion, old_conversion in zip(conversions, self.conversions):
			old_conversion["destination"].set(conversion["destination"])
			old_conversion["source"].set(conversion["source"])
			old_conversion["delimeter"].set(conversion["delimeter"])

	def save_mapping(self, name):
		print "Saving configuration..."
		f = open("mappings.json", "r")
		lines = "".join(f.readlines())
		f.close()

		configs_json = json.loads(lines)
		actual_conversions = []
		for conversion in self.conversions:
			actual_conversions.append({"source":conversion["source"].get(), "destination":conversion["destination"].get(), "delimeter":conversion["delimeter"].get()})

		full_object = {}
		full_object["conversions"] = actual_conversions
		full_object["constants"] = self.constants
		configs_json[name] = full_object

		lines = json.dumps(configs_json)

		f = open("mappings.json", "w+")
		f.write(lines)

	def store_constant(self):
		constant_field = self.constant_field.get()
		constant_val = self.constant_val.get()

		self.constants[constant_field] = constant_val

	def show_constants(self):
		top = Toplevel()
		top.title = "Constants for this mapping"
		top.geometry("400x200+100+100")

		bg = Frame(top)

		for constant in self.constants:
			for_display = "%s: %s" % (constant, self.constants[constant])
			frm = Frame(bg)
			Label(frm, text=for_display).pack(side=LEFT)
			frm.pack(side=TOP, fill=X)

		button = Button(bg, text="Dismiss", command=top.destroy)
		button.pack()

		bg.pack(fill=BOTH, expand=1)




	def remove_mapping(self):
		del self.conversions[-1]
		self.conversion_frames[-1].pack_forget()
		self.conversion_frames[-1].destroy()
		del self.conversion_frames[-1]
		del self.conversion_boxes[-1]
		del self.output_boxes[-1]

	def add_input_file(self):
		print "Adding input file..."
		path = askopenfilename(filetypes=[
			('Excel files', '*.xlsx'),
			('All files', '*')
			],
			parent=self.parent)
		workbook = xlrd.open_workbook(path)
		worksheet = workbook.sheet_by_index(0)
		csvfile = open("temp.csv", 'wb')
		wr = csv.writer(csvfile, quoting=csv.QUOTE_ALL)
		for rownum in xrange(worksheet.nrows):
			wr.writerow(list(x.encode('utf-8') if type(x) == type(u'') else x for x in worksheet.row_values(rownum)))
		csvfile.close()
		csvfile = open("temp.csv", "rU")
		reader = csv.reader(csvfile)
		self.input_column_headers = reader.next()
		csvfile.close()

		current_choices = len(self.conversion_frames)

		if current_choices == 0:
			self.conversions.append({"source":StringVar(), "destination":StringVar(), "delimeter":StringVar()})
			conversion_frame = Frame(self.conversions_frame)
			self.conversion_frames.append(conversion_frame)
			conversion_box = apply(OptionMenu, (conversion_frame, self.conversions[len(self.conversions) - 1]["source"]) + tuple(self.input_column_headers))
			self.conversion_boxes.append(conversion_box)
			conversion_box.pack(side=LEFT)
			conversion_frame.pack(side=TOP, fill=X)
		else:
			for box in self.conversion_boxes:
				box.config(values=self.input_column_headers)


	def add_output_file(self):
		print "Adding output file..."
		path = askopenfilename(filetypes=[
			('csv files', '*.csv'),
			('All files', '*')
			],
			parent=self.parent)
		self.out_path=path
		csvfile = open(path, 'rU')
		reader = csv.reader(csvfile)
		self.output_column_headers = reader.next()
		self.constant_field_selection['menu'].delete(0, 'end')
		for header in self.output_column_headers:
			self.constant_field_selection['menu'].add_command(label=header, command=_setit(self.constant_field, header))
#		self.constant_field_selection = apply(OptionMenu, (self.constants_bar, self.constant_field) + tuple(self.output_column_headers))
#		self.constant_field_selection.pack(side=LEFT)
		csvfile.close()

		current_choices = len(self.output_boxes)

		if current_choices == 0:
			conversion_frame = self.conversion_frames[len(self.conversion_frames) - 1]
			output_box = apply(OptionMenu, (conversion_frame, self.conversions[len(self.conversions) - 1]["destination"]) + tuple(self.output_column_headers))
			self.output_boxes.append(output_box)
			label = Label(conversion_frame, text="  map column to  ")
			label.pack(side=LEFT)
			output_box.pack(side = LEFT)
			delim_label = Label(conversion_frame, text= "  Delimeter (for combined entries):")
			delim_label.pack(side=LEFT)
			delimeter_entry = Entry(conversion_frame, textvar=self.conversions[len(self.conversions) - 1]["delimeter"])
			self.delimeter_entries.append(delimeter_entry)
			delimeter_entry.pack(side=LEFT)
			conversion_frame.pack(side=TOP, fill=X)

		else:
			for box in self.output_boxes:
				box.config(values=self.output_column_headers)

	def generate_conversions(self):
		print "Generating mapping..."
		actual_conversions = {}
		for item in self.conversions:
			print item['destination'].get()
			print item['source'].get()
			actual_conversions[item['destination'].get()] = {'sources':[]}

		for item in self.conversions:
			actual_conversions[item['destination'].get()]['sources'].append(item['source'].get())
			actual_conversions[item['destination'].get()]['delimeter'] = item['delimeter'].get()

		self.reformat_csv("temp.csv", self.out_path, actual_conversions)




	def initUI(self):
		self.parent.title("Excel to csv converter")
		self.conversion_boxes = []
		self.output_boxes = []
		self.delimeter_entries = []
		self.conversion_frames = []
		self.conversions = []
		self.input_column_headers = []
		self.output_column_headers = []
		self.constants = {}
		#self.constants = {'device_model':'SM-N910V', 'default_config_name':'HCR', 'HCREnrollMobileIron/server':'m.mobileiron.net:10261', 'email/device model/device serial':'/SM-N910V/TEMP'}


		f = open("mappings.json", "r")
		lines = "".join(f.readlines())
		f.close()
		configs_json = json.loads(lines)
		saved_configs = []
		for item in configs_json:
			saved_configs.append(item)

		self.load_save_bar = Frame(self)

		load_selection = StringVar()
		self.load_button = Button(self.load_save_bar, text="Load:", command=lambda: self.do_load(load_selection.get()))
		self.load_button.pack(side=LEFT)

		self.load_option_menu = apply(OptionMenu, (self.load_save_bar, load_selection) + tuple(saved_configs))
		self.load_option_menu.pack(side=LEFT)

		Frame(self.load_save_bar, width=2, borderwidth=1, relief=SUNKEN).pack(side=LEFT, fill=Y, padx=5, pady=5)

		save_name = StringVar()
		Label(self.load_save_bar, text="Save mapping as: ").pack(side=LEFT)

		self.save_name_entry = Entry(self.load_save_bar, textvar = save_name)
		self.save_name_entry.pack(side=LEFT)
		self.save_button = Button(self.load_save_bar, text="Save", command=lambda: self.save_mapping(save_name.get()))
		self.save_button.pack(side=LEFT)

		self.load_save_bar.pack(side=TOP)
		self.constants_bar = Frame(self)

		self.constant_field = StringVar()
		self.constant_val = StringVar()
		Label(self.constants_bar, text="Constant value:").pack(side=LEFT)
		self.constant_field_selection = apply(OptionMenu, (self.constants_bar, self.constant_field) + tuple(self.output_column_headers))
		self.constant_field_selection.pack(side=LEFT)
		self.constant_val_entry = Entry(self.constants_bar, textvar=self.constant_val)
		self.constant_val_entry.pack(side=LEFT)
		Button(self.constants_bar, text="Store", command=self.store_constant).pack(side=LEFT)
		Button(self.constants_bar, text="View", command=self.show_constants).pack(side=LEFT)
		self.constants_bar.pack(side=TOP)


		self.conversions_frame = Frame(self, relief=SUNKEN, borderwidth=1)
		self.conversions_frame.pack(fill=BOTH, expand=1)

		self.pack(fill=BOTH, expand=1)

		self.add_excel_button = Button(self, text="Choose Excel file", command=self.add_input_file)
		self.add_excel_button.pack(side=LEFT)

		self.add_csv_button = Button(self, text="Choose csv template", command=self.add_output_file)
		self.add_csv_button.pack(side=LEFT)

		self.add_conversion_button = Button(self, text="Add another mapping", command=self.add_mapping)
		self.add_conversion_button.pack(side=LEFT)

		self.remove_conversion_button = Button(self, text="Delete mapping", command=self.remove_mapping)
		self.remove_conversion_button.pack(side=LEFT)

		self.finish_button = Button(self, text="Generate csv", command=self.generate_conversions)
		self.finish_button.pack(side=RIGHT)

	def reformat_csv(self, csv_in, csv_out, conversions):
		out_headers = []
		with open(csv_out, 'rU') as csvoutfile:
			reader = csv.reader(csvoutfile)
			out_headers = reader.next()
		with open(csv_in, 'rU') as csvfile, open(csv_out, 'wb') as csvoutfile:
			inreader = csv.reader(csvfile)
			outwriter = csv.writer(csvoutfile)
			outwriter.writerow(out_headers)
			headers = inreader.next()
			for row in inreader:
				output = [None]*len(out_headers)
				for out_header in out_headers:
					try:
						output_conversion = conversions[out_header]
						if len(output_conversion['sources']) == 1:
							output[out_headers.index(out_header)] = row[headers.index(output_conversion['sources'][0])]
						else:
							output[out_headers.index(out_header)] = ""
							for source in output_conversion['sources']:
								output[out_headers.index(out_header)] += (row[headers.index(source)] + output_conversion['delimeter'])
								output[out_headers.index(out_header)] = output[out_headers.index(out_header)][:-1*len(output_conversion['delimeter'])]
					except KeyError:
						print "No source given to fill column " + out_header
    			for constant in self.constants:
    				try:
    					index = out_headers.index(constant)
    					if index >= 0 and output[index] != None:
    						output[index] += self.constants[constant]
    					elif index >= 0:
    						output[index] = self.constants[constant]
    				except ValueError:
    					pass
    			outwriter.writerow(output)



def main():
	root = Tk()
	w, h = root.winfo_screenwidth(), root.winfo_screenheight()
	root.geometry("%dx%d+0+0" % (w, h))
	root.update()
	app = ExcelConverterDialog(root)
	root.mainloop()

if __name__ == '__main__':
	main()