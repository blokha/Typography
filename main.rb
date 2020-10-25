require 'gtk3'
require 'win32ole'
require 'Prawn'
require "prawn/measurement_extensions"

def update_order (connection,liststore,liststore2,zakaz)
  recordset = WIN32OLE.new('ADODB.Recordset')
  sql = "
select Заявка_бланки.Наименование, Заявка_бланки.Кол_во, Форматы.Формат,
Материал.Бумага,
Переплет,
Заявка_бланки.Кол_Листов,
Заявка_бланки.Цена,
Форматы.Делитель
from Заявка_бланки,Форматы, Материал
where Заявка_бланки.№_заказа=#{zakaz}
and  Заявка_бланки.Формат=Форматы.key
and Материал.Код=Заявка_бланки.Бумага
"
  recordset.Open(sql, connection)
  liststore.clear
  liststore2.clear
  paper = Hash.new()
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
    if row[5]
      iter[5]=row[5]
      iter[6]=row[1]*row[5]/row[7]
else
iter[6]=row[1]/row[7]
    end
    if paper.has_key?(row[3])
      paper[row[3]]+=iter[6]
    else
      paper[row[3]]=iter[6]
    end
    if row[6]
      iter[7]=row[6]
      iter[8]=(row[6].sub(",",".").to_f*row[1]).to_s
    end
  }

  paper.each_pair{|key,value|
    iter=liststore2.append()
    iter[0]=key
    iter[1]=value
  }
  paper.clear
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

columns_paper = ["Вид бум.","Кол-во А3"]
list_paper=Gtk::ListStore.new(String,Integer);
grid_paper=Gtk::TreeView.new(list_paper)
grid_paper.override_font(text_f2)
# columns_paper_width = [80,80]
(0...columns_paper.size).each{|i|
  cell=Gtk::CellRendererText.new();
  cell.xalign=0.5 unless i==0
  cell.yalign=0.5
  # cell.set_wrap_width columns_paper_width[i]
  cell.set_padding(5, 5)
  col=Gtk::TreeViewColumn.new(columns_paper[i],cell,:text=>i);
  # col.fixed_width = columns_order_width[i]
  col.resizable=true;
  col.set_alignment(1.0)
  col.set_alignment(1.0)
  grid_paper.append_column(col);
}


#order information
columns_order = ["Наименование","Кол-во","Формат","Вид бум.","Переплет","Листов","Кол-во А3","Цена", "Сумма"]
list_order=Gtk::ListStore.new(String,Integer,String,String,String,Integer,Integer,String,String);
grid_order=Gtk::TreeView.new(list_order)
# grid_order.set_enable_grid_lines(true)
grid_order.grid_lines=3
grid_order.columns_autosize
grid_order.override_font(text_f2)
columns_order_width = [365,75,85,85,85,85,85,80,80]
(0...columns_order.size).each{|i|
  cell=Gtk::CellRendererText.new();
  cell.xalign=0.5 unless i==0
  cell.yalign=0.5
  cell.set_wrap_width columns_order_width[i]
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

# grid_zakaz.signal_connect('row-activated') { |treeview,sel_path,column|
#   model = treeview.model
#   path = sel_path
#   iter = model.get_iter(path)
#   update_order(connection,list_order,list_paper,iter[0])
# }

select1=grid_zakaz.selection
select1.signal_connect("changed"){|treeselection|
# list_order.clear
update_order(connection,list_order,list_paper,list_zakaz.get_value(treeselection.selected,0))
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
button_print_stickers.signal_connect("clicked"){



  path = 'C:\Program Files\Tracker Software\PDF Editor\PDFXEdit.exe'
  pdf = Prawn::Document.new(:page_size => "A4",:margin => 0.mm)
  pdf.font_size 14
  pdf.font_families.update(
    "Comic" => {
      :normal      => { :file => 'COMIC.TTF', :font => "Comic" },
      :italic      => { :file => 'COMICI.TTF', :font => "Comic-Italic" },
      :bold        => { :file => 'COMICBD.TTF', :font => "Comic-Bold" },
      :bold_italic => { :file => 'COMICZ.TTF', :font => "Comic-BoldItalic" }
    }
  )
  pdf.font "Comic"
  text_x=10.mm
  text_y=297.mm
  i=1
  select1=grid_zakaz.selection
  iter=select1.selected
  post_text=list_zakaz.get_value(iter,2)
  schet_text="Счет №"+list_zakaz.get_value(iter,3).to_s if list_zakaz.get_value(iter,3).to_i>0
  info_text="(056) 785-08-90"

  list_order.each { |model, path, row|
    blank_text=row[0]
    count_text=row[1]
    pdf.formatted_text_box [
      {:text=>"Получатель \n", :styles => [:bold]},
      {:text =>"#{post_text}\n" },
      {:text=>"Наименование\n", :styles => [:bold]},
      {:text =>"#{blank_text}\n" },
      {:text=>"Кол-во\n", :styles => [:bold]},
      {:text =>"#{count_text}\n" },
      {:text =>"#{info_text}\n" },
      {:text=>"#{schet_text}\n"},
    ] ,
    :at => [text_x,text_y],
    :width => 85.mm,
    :height => 99.mm,
    :align => :center,
    :valign => :center,
    :overflow => :shrink_to_fit

    i=i+1
    text_x=115.mm

    if (i.odd?)
      text_x=0.mm
      text_y=text_y-99.mm
    end
    if i==7
      i=1
      text_x=10.mm
      text_y=297.mm
      pdf.start_new_page
    end
  }
  pdf.render_file "Stickers.pdf"
  Process.spawn(path,"Stickers.pdf")
}

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
