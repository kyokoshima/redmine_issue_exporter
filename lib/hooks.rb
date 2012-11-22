#coding:utf-8
module IssueExporter
	class Hooks < Redmine::Hook::ViewListener
		def view_issues_index_bottom(context={})
			# render_on :view_issues_index_bottom, :partial => 'link'
			out = <<-EOS
				<p class="other-formats">
					Excel出力
					<span> #{link_to('出力', 'issue_export')} </span>
				</p>
			EOS
		end
		
	end
end