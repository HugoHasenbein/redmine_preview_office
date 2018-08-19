# encoding: utf-8
#
# Redmine plugin to preview a Microsoft Office attachment file
#
# Copyright © 2018 Stephan Wenzel <stephan.wenzel@drwpatent.de>
#
# This program is free software; you can redistribute it and/or
# modify it under the terms of the GNU General Public License
# as published by the Free Software Foundation; either version 2
# of the License, or (at your option) any later version.
#
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.
#
# You should have received a copy of the GNU General Public License
# along with this program; if not, write to the Free Software
# Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301, USA.
#

require 'redmine'

Redmine::Plugin.register :redmine_preview_office do
  name 'Redmine Preview Office'
  author 'Stephan Wenzel'
  description 'This is a plugin for Redmine to preview a Microsoft Office attachment file as pdf'
  version '1.0.3'
  url 'https://github.com/HugoHasenbein/redmine_preview_office'
  author_url 'https://github.com/HugoHasenbein/redmine_preview_office'

  # link to settings page
  # pdf_embedding
  #   0 - <object>-tag, fallback with <embed>-tag
  #   1 - <iframe>-tag
  #
  settings :default => {'pdf_embedding' 	=> '0' # use <object><embed>-tag
                       },
           :partial => 'settings/redmine_preview_office/settings'

end

require "redmine_preview_office"
