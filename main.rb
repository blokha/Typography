require 'gtk3'
require 'win32ole'

def update_order (connection,liststore,zakaz) 
recordset = WIN32OLE.new('ADODB.Recordset')
sql = "
select Заявка_бланки.Наименование, Заявка_бланки.Кол_во, Форматы.Формат,
Материал.Бумага,
Переплет,
Заявка_бланки.Кол_Листов
from Заявка_бланки,Форматы, Материал
where Заявка_бланки.№_заказа=#{zakaz}
and  Заявка_бланки.Формат=Форматы.key
and Материал.Код=Заявка_бланки.Бумага
"
recordset.Open(sql, connection)
liststore.clear
return 1 if recordset.EOF
rows=recordset.GetRows.transpose
recordset.close
rows.each { |row|
iter=liststore.append()
iter[0]=row[0]
iter[1]=row[1]
iter[2]=row[2]
iter[3]=row[3]
iter[4]="мягкий" if row[4]==1
iter[4]="твердый" if row[4]==2
iter[5]=row[5] if row[5]
}
end


def update_zakaz (connection,liststore,ftext='',status = true) 
recordset = WIN32OLE.new('ADODB.Recordset')
sql = "
select Заказы.Код, Дата, Заказчики.Сокращенное_название, Счет, Сумма
from Заказы, Заказчики
where Заказы.Заказчик=Заказчики.Код 
and Статус=#{status}
and (Заказчики.Сокращенное_название like '%#{ftext}%'
or Заказы.Код like '%#{ftext}%')
order by Дата  desc
"
recordset.Open(sql, connection)
liststore.clear
return 1 if recordset.EOF
rows=recordset.GetRows.transpose
recordset.close
rows.each { |row|
iter=liststore.append()
iter[0]=row[0]
iter[1]=row[1].strftime("%d/%m/%y")
iter[2]=row[2].to_s
iter[3]=row[3].to_s
iter[4]=row[4].to_s
}
end




Gtk.init


connection = WIN32OLE.new('ADODB.Connection')
# connection.Open('Provider=Microsoft.Jet.OLEDB.4.0;Data Source=1.mdb')
connection.Open('Provider=Microsoft.ACE.OLEDB.12.0;Data Source=1.mdb')


text_f1=Pango::FontDescription.new("Normal bold 12")
#grid text
text_f2=Pango::FontDescription.new("Normal  12")

window = Gtk::Window.new()
window.set_default_size 880,590

window.override_background_color('normal',"#323c4e")
window.override_color('normal',"#3AD900")
window.signal_connect("destroy") { Gtk.main_quit }
window.set_title 'Typography'
window.position='center'
window.resizable=false
window.border_width=15

#zakaz
columns_zakaz = ["№","Дата","Заказчик","Счет","Сумма"]
columns_zakaz_width = [50,100,300,100,100] 
list_zakaz=Gtk::ListStore.new(Integer,String,String,String,String);
grid_zakaz=Gtk::TreeView.new(list_zakaz)
grid_zakaz.override_font(text_f2)

(0...columns_zakaz.size).each{|i|
cell=Gtk::CellRendererText.new();
col=Gtk::TreeViewColumn.new(columns_zakaz[i],cell,:text=>i);
# col.resizable=true;
col.set_sizing('FIXED')
col.fixed_width = columns_zakaz_width[i]
grid_zakaz.append_column(col);
}
#paper information

columns_paper = ["Paper","Count"]
list_paper=Gtk::ListStore.new(String,Integer);
grid_paper=Gtk::TreeView.new(list_paper)
grid_paper.override_font(text_f2)

#order information
columns_order = ["Наименование","Кол-во","Формат","Бумага","Переплет","Листов"]
list_order=Gtk::ListStore.new(String,Integer,String,String,String,Integer);
grid_order=Gtk::TreeView.new(list_order)
# grid_order.set_enable_grid_lines(true)
grid_order.grid_lines=2
grid_order.columns_autosize
grid_order.override_font(text_f2)
columns_order_width = [450,70,80,80,85,60] 
(0...columns_order.size).each{|i|
cell=Gtk::CellRendererText.new();
cell.xalign=0.5 unless i==0
cell.yalign=0.5
cell.set_wrap_width 450
cell.set_wrap_mode :word

cell.set_padding(5, 5)
col=Gtk::TreeViewColumn.new(columns_order[i],cell,:text=>i);

col.fixed_width = columns_order_width[i]
col.resizable=true;
col.set_alignment(1.0)
col.set_alignment(1.0)
grid_order.append_column(col);
}

#scrolwin zakazi
scrollwindow_zakaz=Gtk::ScrolledWindow.new();
scrollwindow_zakaz.set_min_content_height(200);
scrollwindow_zakaz.set_min_content_width(650);
scrollwindow_zakaz.set_policy('automatic','automatic');
scrollwindow_zakaz.add(grid_zakaz);

#scrolwin paper
scrollwindow_paper=Gtk::ScrolledWindow.new();
scrollwindow_paper.set_min_content_height(200);
scrollwindow_paper.set_min_content_width(190);
scrollwindow_paper.set_policy('automatic','automatic');
scrollwindow_paper.add(grid_paper);


#scrolwin zakaz
scrollwindow_order=Gtk::ScrolledWindow.new();
scrollwindow_order.set_min_content_height(300);
scrollwindow_order.set_min_content_width(190);
scrollwindow_order.set_policy('automatic','automatic');
scrollwindow_order.add(grid_order);

grid_zakaz.signal_connect('row-activated') { |treeview,sel_path,column|
model = treeview.model
path = sel_path
iter = model.get_iter(path)
update_order(connection,list_order,iter[0])
}

check_status = Gtk::CheckButton.new()
find = Gtk::Entry.new()

check_status.set_label('В работе')
check_status.set_active(true)
check_status.signal_connect("toggled"){
	update_zakaz(connection,list_zakaz,find.text,check_status.active?)
}

find = Gtk::Entry.new()
find.signal_connect("activate"){
	update_zakaz(connection,list_zakaz,find.text,check_status.active?)
}

button_new = Gtk::Button.new(:label => 'New')
button_open = Gtk::Button.new(:label => 'Open')
button_del = Gtk::Button.new(:label => 'Delete')


hbox1 = Gtk::Box.new('horizontal', 15)
hbox1.pack_start(find,:expand => true, :fill =>true)
hbox1.pack_start(check_status)
hbox1.pack_end(button_new)
hbox1.pack_end(button_open)
hbox1.pack_end(button_del)

button_print_zakaz = Gtk::Button.new(:label => 'Zakaz')
button_print_stickers = Gtk::Button.new(:label => 'Stickers')
button_print_order = Gtk::Button.new(:label => 'Order')


hbox2 = Gtk::Box.new('horizontal', 15)
hbox2.pack_end(button_print_zakaz)
hbox2.pack_end(button_print_stickers)
hbox2.pack_end(button_print_order)



grid = Gtk::Grid.new()
grid.row_spacing = 15
grid.column_spacing = 15
grid.column_homogeneous = true
grid.attach(hbox1,0,0,4,1)
grid.attach(scrollwindow_zakaz,0,1,3,1)
grid.attach(scrollwindow_paper,3,1,1,1)
grid.attach(scrollwindow_order,0,2,4,1)
grid.attach(hbox2,0,3,4,1)
window.add(grid)
update_zakaz(connection,list_zakaz)
window.show_all
Gtk.main