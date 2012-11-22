# Plugin's routes
# See: http://guides.rubyonrails.org/routing.html

match 'projects/:project_id/issue_export', :to => 'issue_export#index'
