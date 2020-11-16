require 'gtk3'
require 'win32ole'
require 'Prawn'
require 'prawn/table'
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
      iter[7]=row[6].sub(",",".")
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
connection.Open('Provider=Microsoft.ACE.OLEDB.12.0;Data Source=Изготовление продукции.mdb')


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

button_print_zakaz = Gtk::Button.new(:label => 'Заказ')
button_print_stickers = Gtk::Button.new(:label => 'Этикетки')
button_print_order = Gtk::Button.new(:label => 'Счет')
#Print_order
button_print_order.signal_connect ("clicked"){
  path_pdf = 'C:\Program Files\Tracker Software\PDF Editor\PDFXEdit.exe'
  pdf = Prawn::Document.new(:page_size => "A4",:margin => 20.mm)
  pdf.font_size 14
  pdf.font_families.update(
    "Comic" => {
      :normal      => { :file => 'COMIC.TTF', :font => "Comic" },
      :italic      => { :file => 'COMICI.TTF', :font => "Comic-Italic" },
      :bold        => { :file => 'COMICBD.TTF', :font => "Comic-Bold" },
      :bold_italic => { :file => 'COMICZ.TTF', :font => "Comic-BoldItalic" }
  })
  pdf.font "Comic"

  select1=grid_zakaz.selection
  iter=select1.selected
  post_1=list_zakaz.get_value(iter,2)
  post_2=list_zakaz.get_value(iter,1)
  pdf.text post_1+"  "+post_2,:align => :center, :style => :bold
  pdf.move_down 10
  data = Array.new {Array.new(4)}
  data.push ["Наименование","Кол-во","Цена","Сумма"]
  sum = 0
  list_order.each { |model, path, row|
    data.push [row[0],row[1],sprintf("%.2f" % row[7].to_f),sprintf("%0.2f" % row[8].to_f)]
    sum = sum+row[8].to_f
  }
  data.push ["","","Всего",sprintf("%0.2f" % sum)]
  pdf.table data do
    row(0).font_style = :bold
    row(0).align = :center
    column(0).width = 95.mm
    column(1).width = 25.mm
    column(2).width = 25.mm
    column(3).width = 25.mm
    column(1..3).align = :center
    column(1..3).valign = :center
    row(-1).font_style = :bold
    row(-1).column(0..3).borders = [:top]
  end
  pdf.render_file "Price.pdf"
  Process.spawn(path_pdf,"Price.pdf")

}

#Print stickers
button_print_stickers.signal_connect("clicked"){
  path_pdf = 'C:\Program Files\Tracker Software\PDF Editor\PDFXEdit.exe'
  pdf = Prawn::Document.new(:page_size => "A4",:margin => 0.mm)

  pdf.line 105.mm,0,105.mm,297.mm,0
  pdf.line 0,99.mm,210.mm,99.mm,0
  pdf.line 0,198.mm,210.mm,198.mm,0
  pdf.stroke

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
  rows_my = Array.new() {Array.new(2)}
  # dialog = Gtk::MessageDialog.new(
  #   :parent => window,
  #   :type => :question,
  # :buttons => Gtk::ButtonsType::YES_NO)
  dialog = Gtk::Dialog.new(
    # :parent => window,
    :flag => :modal
  )
  dialog.signal_connect('destroy') { dialog.destroy }
  dialog.set_title "Разбить?"
  dialog.position="center"
  box = dialog.content_area()
  dialog_label = Gtk::Label.new()
  dialog_label.justify = :center
  yes_button = dialog.add_button(Gtk::Stock::YES, :yes)
  no_button = dialog.add_button(Gtk::Stock::NO, :no)
  apply_button = dialog.add_button(Gtk::Stock::NEW, :accept)
  box.add(dialog_label)
  dialog.show_all
  list_order.each { |model, path, row|
    if row[6]>1000
      m = row[6].to_i / 1000
      if (row[6].to_i % 1000)>0 then m=m+1 end
      count = row[1].to_i / m
      count1 = row[1].to_i - (m-1)*count
      if count == count1
      str1 = "на #{m} x #{count}"
	  else
      str1 = "на #{m-1} x #{count} \nи 1 x #{count1}"
	  end
      dialog_label.set_markup("<span size='14'>Разбить\n<b>#{row[0]}</b>\nв количестве #{row[1]}\n#{str1}?\n</span>")

      response = dialog.run
      if response == :yes
        (1..m-1).each { rows_my.push [row[0],count] }
        rows_my.push [row[0],count1] if count1>0
        next
      end
      if response == :accept
        dialog1 = Gtk::Dialog.new(
          :parent => dialog,
          :flag => :modal
        )
        dialog1.signal_connect('destroy') { dialog1.destroy }
        dialog1.set_title "Разбить?"
        dialog1.position="center"
        box1 = dialog1.content_area()
        dialog1_label = Gtk::Label.new()
        dialog1_label_1 = Gtk::Label.new()
        dialog1_label_2 = Gtk::Label.new()
        dialog1_label.set_markup("<span size='14'>Введите данные вручную</span>")
        dialog1_label_1.set_markup("<span size='14'> по </span>")
        dialog1_label_2.set_markup("<span size='14'> по </span>")
        dialog1_entry_m1 = Gtk::Entry.new()
        dialog1_entry_m1.max_length = 2
        dialog1_entry_m1.width_chars = 2
        dialog1_entry_m1.text = (m-1).to_s
        dialog1_entry_m2 = Gtk::Entry.new()
        dialog1_entry_m2.max_length = 2
        dialog1_entry_m2.width_chars = 2
        dialog1_entry_m2.text = '1'
        dialog1_entry_count1 = Gtk::Entry.new()
        dialog1_entry_count1.max_length = 4
        dialog1_entry_count1.width_chars = 4
        dialog1_entry_count1.text = count.to_s
        dialog1_entry_count2 = Gtk::Entry.new()
        dialog1_entry_count2.max_length = 4
        dialog1_entry_count2.width_chars = 4
        dialog1_entry_count2.text = count1.to_s
        yes_button1 = dialog1.add_button(Gtk::Stock::OK, :ok)
        no_button1 = dialog1.add_button(Gtk::Stock::CANCEL, :cancel)
        hbox1 = Gtk::Box.new('horizontal',15)
        hbox1.pack_start(dialog1_entry_m1)
        hbox1.pack_start(dialog1_label_1)
        hbox1.pack_start(dialog1_entry_count1)
        hbox2 = Gtk::Box.new('horizontal',15)
        hbox2.pack_start(dialog1_entry_m2)
        hbox2.pack_start(dialog1_label_2)
        hbox2.pack_start(dialog1_entry_count2)
        box1.add(dialog1_label)
        box1.add(hbox1)
        box1.add(hbox2)
        dialog1.show_all
        response = dialog1.run
        if response == :ok
        	(1..dialog1_entry_m1.text.to_i).each { rows_my.push [row[0],dialog1_entry_count1.text.to_i] }
        	(1..dialog1_entry_m2.text.to_i).each { rows_my.push [row[0],dialog1_entry_count2.text.to_i] }
	        dialog1.destroy
    	next
	    else
	    	dialog1.destroy
	    end
        
        
      end
    end

    rows_my.push [row[0],row[1]]
  }
  dialog.destroy

  rows_my.each{ |row|
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
    :width => 75.mm,
    :height => 99.mm,
    :align => :center,
    :valign => :center,
    :overflow => :shrink_to_fit

    i=i+1
    text_x=120.mm

    if (i.odd?)
      text_x=15.mm
      text_y=text_y-99.mm
    end
    if i==7
      i=1
      text_x=15.mm
      text_y=297.mm
      pdf.start_new_page
      pdf.line 105.mm,0,105.mm,297.mm,0
      pdf.line 0,99.mm,210.mm,99.mm,0
      pdf.line 0,198.mm,210.mm,198.mm,0
      pdf.stroke
    end
  }
  pdf.render_file "Stickers.pdf"
  Process.spawn(path_pdf,"Stickers.pdf")
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
