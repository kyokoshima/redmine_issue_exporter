#coding: utf-8
require_dependency 'spreadsheet'
module IssueExportHelper

	def create_excel_book
      cached_pos = {}
			project = Project.find(params[:project_id], 
  			:include => {:issues => [:tracker, :author, :status, :priority, :category, :fixed_version]})
    	# book = Spreadsheet::Workbook.new
      tpl_book = Spreadsheet.open "#{Rails.root}/plugins/redmine_issue_exporter/test/tpl1.xls"
      tpl_sheet = tpl_book.worksheet(0).dup
      tpl_book.io.close
      download_book = Spreadsheet::Workbook.new
    	project.issues.each do |issue|
    		s= download_book.create_worksheet(:name=>"##{issue.id} - #{create_xls_sheet_name(issue.subject)}")
        tpl_sheet.each do |row|
          row.each_with_index do |cell, col_index|
            cell_value = cell.to_s
            if cell_value.start_with?('$')
              key = cell_value.gsub(/\$/, '')
              logger.debug("key ----- #{key}")
              begin
                cell_value = eval(key)
              rescue NameError
                logger.debug("not found ----- #{key}")
                ''
              end
            end
            s[row.idx, col_index] = cell_value
            s.row(row.idx).set_format(col_index, row.format(col_index))
            s.format_column(col_index, row.format(col_index), {:width=>tpl_sheet.column(col_index).width})
             # s.format_column(col_index, row.format(col_index), {:width=>3})
          end
        end
        
        # s[0,0] = "#{project.name} - #{i.tracker.name} #{i.id}"
    		# s[1,0] = "#{i.subject}"
    		# s[2,0] = "#{format_time(i.created_on)} - #{i.author.lastname}#{i.author.firstname}"

    		# s[4,0] = "#{l(:field_status)}: #{i.status.name}"
    		# s[5,0] = "#{l(:field_priority)}: #{i.priority.name}"
    		# s[6,0] = "#{l(:field_assigned_to)}: #{i.author.name}"
    		# s[7,0] = "#{l(:field_category)}: #{i.category.name if i.category.present?}"
    		# s[8,0] = "#{l(:field_fixed_version)}: #{i.fixed_version.name if i.fixed_version.present?}"
    		 
    		# s[4,10] = "#{l(:field_start_date)}: #{i.start_date}"
    		# s[5,10] = "#{l(:field_due_date)}: #{i.due_date}"
    		# s[6,10] = "#{l(:field_done_ratio)}: #{i.done_ratio}"
    		# s[7,10] = "#{l(:field_estimated_hours)}: #{i.estimated_hours}"
    		# s[8,10] = "#{l(:label_spent_time)}: #{i.total_spent_hours}"
    		# s[10,0] = "#{l(:field_description)}"
    		# s[11,0] = "#{i.description}"

      #   s[12,0] = "#{l(:label_history)}"
      #   history_row = 13
      #   history_index = 1
      #   i.journals.find(:all, :include=>[:user, :details], :order=>"#{Journal.table_name}.created_on asc").each do |j|
      #     s[history_row, 0] = ["##{history_index}","#{format_time(j.created_on)}","#{j.user.name}"].join(" - ")
      #     logger.debug("details --- #{j.details}")
      #     details_to_strings(j.details, true).each do |d|
      #       # s[history_row+=1, 0] = "- #{j.details}"
      #       s[history_row+=1, 0] = "- #{d}"
      #   end
      #   if j.notes?
      #     s[history_row+=1, 0] = "#{j.notes}"
      #   end

      #   if i.attachments.any?
      #     s[history_row+=1, 0] = "#{l(:label_attachment_plural)}"
      #     i.attachments.each do |a|
      #       s[history_row+=1, 0] = ["#{a.filename}","#{number_to_human_size(a.filesize)}",
      #         "#{format_date(a.created_on)}","#{a.author.name}"].join(',')
      #     end
      #   end
      #   history_row += 1
      #   history_index += 1
      # end
  	end
  	# sheet1[0,0] = @project
    data = StringIO.new
    download_book.write data
    # download_book.io.close
    data.string
	end
	def create_xls_sheet_name(src_name)
    #:\?[]/*の全角半角を|に変換
    name = src_name.gsub(/[:\\\?\[\]\/\*：￥？／＊]/,"|")

    #31文字以上の場合31文字目を…にする
    name = name[0, 30] + "…" if name.slice(//).length > 30
    logger.debug("sheet name is === #{name}")
    name
	end
end