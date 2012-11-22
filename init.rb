require_dependency 'hooks.rb'
Redmine::Plugin.register :issue_exporter do
  name 'Issue Exporter plugin'
  author 'Author name'
  description 'This is a plugin for Redmine'
  version '0.0.1'
  url 'http://example.com/path/to/plugin'
  author_url 'http://example.com/about'

	project_module :issue_exporter do
			permission :view_issue_exporter, :index => [:index]
	end
end

