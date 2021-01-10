require 'gtk3'
require 'win32ole'



def write_to_list_from_sql (connection,liststore,tablename)
  ####
  recordset = WIN32OLE.new('ADODB.Recordset')
  sql = "select * from #{tablename}
order by 1"
  recordset.Open(sql, connection)
  rows = recordset.GetRows.transpose
  recordset.close
  return 1 if rows.empty?
  iter = liststore.append
    iter[0] = 0
    iter[1] = "-"
  rows.each {|row|
    iter = liststore.append
    iter[0] = row[0]
    iter[1] = row[1]
  }

end

Gtk.init
# connection = WIN32OLE.new('ADODB.Connection')
# connection.Open('Provider=Microsoft.ACE.OLEDB.12.0;Data Source=Изготовление продукции.mdb')

text_f1=Pango::FontDescription.new("Normal bold 12")
#grid text
text_f2=Pango::FontDescription.new("Normal  12")

id_zakaz = -1

connection = WIN32OLE.new('ADODB.Connection')
# connection.Open('Provider=Microsoft.Jet.OLEDB.4.0;Data Source=1.mdb')
connection.Open('Provider=Microsoft.ACE.OLEDB.12.0;Data Source=Изготовление продукции.mdb')


window = Gtk::Window.new()
window.set_default_size 1100,590
window.override_background_color('normal',"#323c4e")
window.override_color('normal',"#3AD900")
window.signal_connect("destroy") { Gtk.main_quit }
window.set_title 'Typography'
window.position='center'
window.resizable=false
window.border_width=15
#zakazchik information
column_zakazchik = ["Полное название","Сокращенное название","Адрес","Телефон","Код"]
list_zakazchik = Gtk::ListStore.new(String,String,String,String,Integer)
grid_zakazchik = Gtk::TreeView.new(list_zakazchik)
grid_zakazchik.grid_lines=3
grid_zakazchik.columns_autosize
grid_zakazchik.override_font(text_f2)
columns_zakazchik_width = [450,300,150,100]
(0...columns_zakazchik_width.size).each{|i|
  cell = Gtk::CellRendererText.new()
  cell.xalign=0.5 unless i==0
  cell.yalign=0.5
  cell.set_wrap_width(columns_zakazchik_width[i])
  cell.set_wrap_mode :word
  cell.set_padding(5, 5)
  col=Gtk::TreeViewColumn.new(column_zakazchik[i],cell,:text=>i);

  col.fixed_width = columns_zakazchik_width[i]
  col.resizable=true;
  col.set_alignment(0.5)
  grid_zakazchik.append_column(col)
}

#scrolwin zakazhcik
scrollwindow_zakazchik=Gtk::ScrolledWindow.new();
scrollwindow_zakazchik.set_min_content_height(150);
scrollwindow_zakazchik.set_min_content_width(190);
scrollwindow_zakazchik.set_policy('automatic','automatic');
scrollwindow_zakazchik.add(grid_zakazchik);


#order grid

columns_order = ["Наименование","Кол-во","Формат","Вид бум.","Переплет","Листов","Шнур","Нум","Цена"]
list_order=Gtk::ListStore.new(String,Integer,String,String,String,Integer,Integer,Integer,String);
grid_order=Gtk::TreeView.new(list_order)
grid_order.grid_lines=3
grid_order.columns_autosize
grid_order.override_font(text_f2)
columns_order_width = [365,75,85,85,85,85,85,80,80]
(0...columns_order.size).each{|i|
  cell=Gtk::CellRendererText.new()
  cell.xalign=0.5 unless i==0
  cell.yalign=0.5
  cell.set_wrap_width columns_order_width[i]
  cell.set_wrap_mode :word
  cell.set_padding(5, 5)
  col=Gtk::TreeViewColumn.new(columns_order[i],cell,:text=>i);
  col.fixed_width = columns_order_width[i]
  col.resizable=true;
  col.set_alignment(0.5)
  grid_order.append_column(col)
}



#scrolwin order
scrollwindow_order=Gtk::ScrolledWindow.new();
scrollwindow_order.set_min_content_height(350);
scrollwindow_order.set_min_content_width(190);
scrollwindow_order.set_policy('automatic','automatic');
scrollwindow_order.add(grid_order);

#order grid find
columns_order_find = ["Наименование","Формат","Вид бум.","Переплет","Листов","Шнур","Нум","Цена"]
columns_order_find_width = [665,100,100,85,85,85,80,80]
list_order_find=Gtk::ListStore.new(String,String,String,String,String,String,String,String);
grid_order_find=Gtk::TreeView.new(list_order_find)
grid_order_find.grid_lines=3
grid_order_find.columns_autosize
grid_order_find.override_font(text_f2)
(0...columns_order_find.size).each{|i|
  cell=Gtk::CellRendererText.new()
  cell.xalign=0.5 unless i==0
  cell.yalign=0.5
  cell.set_wrap_width columns_order_find_width[i]
  cell.set_wrap_mode :word
  cell.set_padding(5, 5)
  col=Gtk::TreeViewColumn.new(columns_order_find[i],cell,:text=>i);
  col.fixed_width = columns_order_find_width[i]
  col.resizable=true;
  col.set_alignment(0.5)
  grid_order_find.append_column(col)
}



#scrolwin zakaz find
scrollwindow_order_find=Gtk::ScrolledWindow.new();
scrollwindow_order_find.set_min_content_height(150);
scrollwindow_order_find.set_min_content_width(190);
scrollwindow_order_find.set_policy('automatic','automatic');
scrollwindow_order_find.add(grid_order_find);

zakazchik = Gtk::Entry.new()
zakazchik.signal_connect("changed") {
  recordset = WIN32OLE.new('ADODB.Recordset')
  sql = "select Полное_название,Сокращенное_название, Адрес, Код from Заказчики
where Полное_название like '%#{zakazchik.text}%'
or Сокращенное_название like '%#{zakazchik.text}%'
"
  recordset.Open(sql, connection)
  next if recordset.EOF
  rows = recordset.GetRows.transpose
  recordset.close
  list_zakazchik.clear
  next if rows.empty?
  rows.each {|row|
    iter = list_zakazchik.append
    iter[0] = row[0]
    iter[1] = row[1]
    iter[2] = row[2]
    iter[4] = row[3]
  }
}
label_z = Gtk::Label.new("NONE")
button_add = Gtk::Button.new(:label => 'Create')
button_add.signal_connect("clicked"){
  selection = grid_zakazchik.selection
  next unless selection.selected
  id_zakaz = list_zakazchik.get_value(selection.selected,4)
  label_z.text = list_zakazchik.get_value(selection.selected,1)
}



hbox1 = Gtk::Box.new('vertical', 15)
hbox1.pack_start(zakazchik,:expand => false)
hbox1.pack_start(button_add)
hbox1.pack_end(label_z)



format_list = Gtk::ListStore.new(Integer,String)
paper_list = Gtk::ListStore.new(Integer,String)
nym_lisr = Gtk::ListStore.new(Integer,String)
iter = nym_lisr.append; iter[0] = 1 ;iter[1]="-"
iter = nym_lisr.append; iter[0] = 2 ;iter[1]="шнуровать"
iter = nym_lisr.append; iter[0] = 3 ;iter[1]="нумеровать"
iter = nym_lisr.append; iter[0] = 4 ;iter[1]="шнур, нумер"
pereplet_list = Gtk::ListStore.new(Integer,String)
write_to_list_from_sql(connection,format_list,"Форматы")
write_to_list_from_sql(connection,paper_list,"Материал")
write_to_list_from_sql(connection,pereplet_list,"Переплет")



entry_name = Gtk::Entry.new
entry_kol = Gtk::Entry.new
entry_listov = Gtk::Entry.new


entry_format = Gtk::ComboBox.new(:entry => true, :model =>format_list, :area=>nil)
entry_format.set_entry_text_column(1)
entry_format.active = 3

entry_paper = Gtk::ComboBox.new(:entry => true, :model =>paper_list, :area=>nil)
entry_paper.set_entry_text_column(1)
entry_paper.active = 2

entry_nym = Gtk::ComboBox.new(:entry => true, :model =>nym_lisr, :area=>nil)
entry_nym.set_entry_text_column(1)

entry_pereplet = Gtk::ComboBox.new(:entry => true, :model =>pereplet_list, :area=>nil)
entry_pereplet.set_entry_text_column(1)




#
entry_name.width_chars = 55
entry_kol.width_chars = 5
entry_listov.width_chars = 3
entry_listov.max_length = 3
label1 = Gtk::Label.new()
label1.text="Наименование"+" "*150+"Кол-во"+" "*13+"Формат"+" "*58+"Бумага"+" "*59+"Листов"+" "*8+"ШнурНум"+" "*54+"Переплет"
label1.set_xalign (-1)
entry_name.override_font(text_f1)
entry_kol.override_font(text_f1)
entry_format.override_font(text_f1)
entry_paper.override_font(text_f1)
entry_listov.override_font(text_f1)
entry_nym.override_font(text_f1)
entry_pereplet.override_font(text_f1)

entry_name.signal_connect("changed"){
  next unless id_zakaz
  recordset = WIN32OLE.new('ADODB.Recordset')

  sql = "select  distinct(Заявка_бланки.Наименование),
Форматы.Формат,
Материал.Бумага,
Переплет.Название,
Заявка_бланки.Кол_Листов,
Заявка_бланки.Шнурация,
Заявка_бланки.Нумерация,
Заявка_бланки.Цена
from (((Заявка_бланки INNER JOIN Заказы 
ON Заявка_бланки.№_заказа = Заказы.Код) LEFT JOIN Форматы
ON Форматы.key = Заявка_бланки.Формат) LEFT JOIN Материал
ON Материал.Код = Заявка_бланки.Бумага) LEFT JOIN Переплет
ON Переплет.Код = Заявка_бланки.Переплет
where Заказы.Заказчик = #{id_zakaz}
and Заявка_бланки.Наименование like '%#{entry_name.text}%'
order by 1"
  recordset.Open(sql, connection)
  list_order_find.clear
  entry_listov.text = ""
  entry_paper.active = 0
  entry_format.active = 0
  entry_nym.active = 0
  entry_pereplet.active = 0
  next if recordset.EOF
  rows = recordset.GetRows.transpose
  recordset.close

  rows.each {|row|
    iter = list_order_find.append
    # columns_order_find = ["Наименование"0,"Формат"1,"Вид бум."2,"Переплет"3,"Листов"4,"Шнур"5,"Нум"6,"Цена"7]
    iter[0] = row[0]
    iter[1] = row[1].to_s
    iter[2] = row[2].to_s
    iter[3] = row[3].to_s
    iter[4] = row[4].to_s if row[4]
    iter[5] = "х" if row[5]
    iter[6] = "x" if row[6]
    iter[7] = row[7].sub(",",".") if row[7]


  }
}

grid_order_find.signal_connect('row-activated') { |treeview,sel_path,column|
  model = treeview.model
  path = sel_path
  iter = model.get_iter(path)
  entry_name.text = iter[0]
  entry_kol.text = "1"
  format_list.each{|model1,path1,iter1|
    if format_list.get_value(iter1,1) == iter[1]
      entry_format.active = path1.to_s.to_i
    end
  }  
  paper_list.each{|model1,path1,iter1|
    if paper_list.get_value(iter1,1) == iter[2]
      entry_paper.active = path1.to_s.to_i
    end
  }
    pereplet_list.each{|model1,path1,iter1|
    if pereplet_list.get_value(iter1,1) == iter[3]
      entry_pereplet.active = path1.to_s.to_i
    end
  }
  entry_listov.text = iter[4].to_s
if iter[5] and iter[6]
	entry_nym.active = 3
elsif iter[5]
	entry_nym.active = 1
elsif iter[6]
	entry_nym.active = 2
end
}

button_add_row = Gtk::Button.new(:label => "Добавить")
button_save_row = Gtk::Button.new(:label => "Сохранить")
button_delete_row = Gtk::Button.new(:label => "Удалить")

hbox2 = Gtk::Box.new('horizontal', 15)
hbox2.pack_start(entry_name)
hbox2.pack_start(entry_kol)
hbox2.pack_start(entry_format)
hbox2.pack_start(entry_paper)
hbox2.pack_start(entry_listov)
hbox2.pack_start(entry_nym)
hbox2.pack_start(entry_pereplet)

hbox3 = Gtk::Box.new('vertical', 15)
hbox3.pack_start(button_add_row)
hbox3.pack_start(button_save_row)
hbox3.pack_end(button_delete_row)

grid = Gtk::Grid.new()
grid.row_spacing = 15
grid.column_spacing = 15
grid.column_homogeneous = true
grid.attach(hbox1,0,0,1,5)
grid.attach(scrollwindow_zakazchik,1,0,4,5)
grid.attach(label1,0,5,5,1)
grid.attach(hbox2,0,6,5,1)
grid.attach(scrollwindow_order_find,0,7,4,1)
grid.attach(hbox3,4,7,1,1)
grid.attach(scrollwindow_order,0,8,5,1)
window.add(grid)




window.show_all
Gtk.main
