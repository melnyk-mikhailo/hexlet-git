require 'rest_client'
require 'json'
require 'date'
require 'rubyXL'
workbook = RubyXL::Parser.parse '/home/misha/work/land_document.xlsx'
worksheet = workbook[0]

require "prawn"
font_regular = '/home/misha/work/times.ttf'
font_bolt = '/home/misha/work/timesbd.ttf'


#get vkursi token
def get_vkursi_token
  values = '{
    "email":"S",
    "password":""
  }'
  headers = {
    :content_type => 'application/json'
  }
  #puts headers
  r = RestClient.post 'https://vkursi-api.azurewebsites.net/api/1.0/token/authorize', values, headers
  response = JSON.parse(r)
  token = response['token']
  #puts token
  #puts r.code
  token
end
#puts vkursi_token

def ApiConstructorZemli(vkursi_token,cadnumber,task_id)
  headers = {:content_type => 'application/json',:Authorization => "Bearer #{vkursi_token}"}
  #puts headers
  #values = '{"Cadastrs": ["7121589200:05:004:0521"], "taskId":"72e836d8-b559-41d4-8528-db648ce59c8f"}'
  values = {"Cadastrs" =>["#{cadnumber}"],"TaskId" => "#{task_id}","MethodsList"=>[0]}.to_json
  #puts values.class
  #puts values
  r = RestClient.post 'https://vkursi-api.azurewebsites.net/api/2.0/ApiConstructorZemli', values, headers
  puts "vkursi response #{r.code}"
  #puts r
  response = JSON.parse(r)

  #save response to json file
  #puts response
end


def save_json(data)
  File.write("/home/misha/work/data.json",JSON.pretty_generate(data))
end

def convert_date (value,text)
  if value.nil?
    nil
  else
    date_convert = DateTime.strptime(value,'%Y-%m-%dT%H:%M:%S')
    str_date = date_convert.strftime("%d.%m.%Y")
    "\n" + "#{text}" + str_date + "\n"
  end
end

def check_str(value,text)
  if value.nil?
    nil
  else
    "\n" +"#{text}" + value + "\n"
  end
end

def pdf_generate(pdf,text_write,type_font)
  font_regular = '/home/misha/work/times.ttf'
  font_bolt = '/home/misha/work/timesbd.ttf'
  if type_font == 'font_bolt'
    pdf.move_down 10
    pdf.font font_bolt
    pdf.text text_write,size:15, align: :center
  elsif type_font == 'font_min'
    pdf.move_down 5
    pdf.font font_regular
    pdf.text text_write,size:10
  elsif type_font == 'font_medium'
    pdf.move_down 15
    pdf.font font_bolt
    pdf.text text_write,size:13, align: :center
  else 
    pdf.move_down 5
    pdf.font font_regular
    pdf.text text_write,size:12
  end
end


vkursi_token = get_vkursi_token
vkursi_id = '8da87edd-2961-4feb-b1bb-34796c0cf62a'

#loop for sheet
worksheet.drop(1).each { |row|
  val_cad = row[3].value
  cadastr = val_cad if (val_cad != nil)
  pdf = Prawn::Document.new 
  #cadastr = (row[3].value).to_s
  puts cadastr
  #response vkursi
  response = ApiConstructorZemli(vkursi_token,cadastr,vkursi_id)['data']
  response.each do |data|
    save_json(data)
    pdf_generate(pdf,"Відомості про земельну ділянку",'font_bolt')

    #Загальна кадастрова інформація
    cadnumber = check_str((data['cadastr']).to_s,'Кадастровий номер земельної ділянки: ')
    general_info = ''
    data['plotGeneralInfo'].each {|plot| 
      purpose = check_str(plot['landZone'],'Цільове призначення: ')
      category = check_str(plot['category'],'Категорія земельної ділянки: ')
      area = check_str((plot['area']).to_s,'Площа земельної ділянки: ')
      region = check_str(plot['region'],'Область: ')
      district = check_str(plot['district'],'Район: ')
      onm = check_str( (plot['onm']).to_s,"Реєстраційний номер об’єкта нерухомого майна: ")
      general_info = cadnumber + purpose + category + area + region + district + onm
    }
    pdf_generate(pdf,general_info,font_regular)

    #Відомості про суб'єктів
    object_subject = ''
    data['plotOwnershipInfo'].each {|owner|
      ship_type = check_str( owner['ownershipType'],'Тип власності: ')
      operation_reason = check_str(owner['operationReason'], 'Орган, що здійснив державну реєстрацію права (в державному реєстрі прав): ')
      doc_date = convert_date(owner['ownershipDocDate'],'Дата державної реєстрації права (в державному реєстрі прав): ')
      owner_code = owner['ownerCode']
      if (owner_code).empty?
        owner = check_str(owner['ownerName'], "Прізвище, ім'я та по батькові фізичної особи: ")
      else
        owner =  check_str(owner['ownerName'], "Найменування юридичної особи: ")
        owner += check_str(owner_code, "Код ЄДРПОУ юридичної особи")
      end
      object_subject = ship_type + owner + doc_date + operation_reason
    }
    pdf_generate(pdf,"Відомості про суб'єктів права власності на земельну ділянку",'font_medium')
    pdf_generate(pdf,"* інформація про власника (землекористувачів) є довідковою, актуальна інформація міститься у Державному реєстрі речових прав на нерухоме майно",'font_min')
    pdf_generate(pdf,object_subject,font_regular)

    #Відомості про суб'єкта речового права на земельну ділянку
    right_info = ''
    data['plotUseRightInfo'].each {|subject|
      right_type = check_str(subject['rightType'], 'Вид речового права: ')
      right_reg_number = check_str((subject['rightRegNum']).to_s,"Номер запису про право (в державному реєстрі прав): ")
      right_start_date = convert_date(subject['rightStartDate'],"Дата державної реєстрації права (в державному реєстрі прав): ")
      right_end_date = convert_date(subject['rightEndDate'],"Дата закінчення: ")
      right_register = check_str(subject['rightRegister'], "Орган, що здійснив державну реєстрацію права (в державному реєстрі прав): ")
      user_code = subject['userCode']
      if (user_code).nil?
        user_name = check_str(subject['userName'], "Прізвище, ім'я та по батькові фізичної особи: ")
      else
        user_name = check_str(subject['userName'], "Найменування юридичної особи: ")
        user_name += check_str(subject['userCode'], "Код ЄДРПОУ юридичної особи: ")
      end
      right_info = right_type + user_name + right_reg_number + right_start_date + right_end_date + right_register
    }
    pdf_generate(pdf,"Відомості про суб'єкта речового права на земельну ділянку",'font_medium')
    pdf_generate(pdf,right_info,font_regular)

    pdf.render_file "/home/misha/work/#{cadastr}.pdf"
  end
  }

  #ruby /home/misha/work/post-vkursi.rb