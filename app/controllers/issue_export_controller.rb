#coding:utf-8
require_dependency 'spreadsheet'
class IssueExportController < ApplicationController
  unloadable


  def index
  	project = Project.find(params[:project_id], 
  			:include => {:issues => [:tracker, :author, :status, :priority, :category, :fixed_version]})
  	book = Spreadsheet::Workbook.new
  	project.issues.each do |i|
  		s= book.create_worksheet(:name=>"##{i.id} - #{i.subject}");
  		s[0,0] = "#{project.name} - #{i.tracker.name}"
  		s[1,0] = "#{i.subject}"
  		s[2,0] = "#{format_time(i.created_on)} - #{i.author.lastname}#{i.author.firstname}"

  		s[4,0] = "ステータス: #{i.status.name}"
  		s[5,0] = "優先度: #{i.priority.name}"
  		s[6,0] = "担当者: #{i.author.name}"
  		s[7,0] = "カテゴリ: #{i.category.name if i.category.present?}"
  		s[8,0] = "対象バージョン: #{i.fixed_version.name if i.fixed_version.present?}"
  		 
  		s[4,10] = "開始日: #{i.start_date}"
  		s[5,10] = "期日: #{i.due_date}"
  		s[6,10] = "進捗%: #{i.done_ratio}"
  		s[7,10] = "予定工数: #{i.estimated_hours}"
  		s[8,10] = "作業時間の記録: #{i.total_spent_hours}"
  		s[10,0] = "説明"
  		s[11,0] = "#{i.description}"
  	end
  	# sheet1[0,0] = @project
  	data = StringIO.new
  	book.write data
  	send_data(data.string, :filename=>"#{project.name}.xls")


  	# render :text => project
  	# tmp.open
  	# send_data(tmp.open)
  	# tmp.close

  	# render :text => 'index', :layout => false
  	# render :text=> book.methods
  end
end
