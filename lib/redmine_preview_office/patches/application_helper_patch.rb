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

module RedminePreviewOffice
  module Patches 
    module ApplicationHelperPatch
      def self.included(base)
        base.send(:include, InstanceMethods)
        base.class_eval do

          unloadable 
          
          def preview_office_tag(attachment, options={})
          
            case Setting['plugin_redmine_preview_office']['pdf_embedding'].to_i
            
              when 0
                content_tag(:div, 
							 content_tag(
								 :object,
								 tag(:embed, :href => preview_office_path(attachment), :type => "application/pdf", :onload  => "resizeObject(this);"),
								 { :style   => "position:absolute;top:0;left:0;width:95%;height:100%;",
								   :title   => attachment.filename,
								   :type    => "application/pdf",
								   :data    => preview_office_path(attachment),
								   :onload  => "resizeObject(this);"
								  }.merge(options)
							 ),
							 :style => "position:relative;padding-top:141%;"
                )
              else
                content_tag(:div, 
							 content_tag(
								 :iframe,
								 "",
								 { :style                => "position:absolute;top:0;left:0;width:95%;height:100%;",
								   :seamless             => "seamless",
								   :scrolling            => "no",
								   :frameborder          => "0",
								   :allowtransparency    => "true",
								   :title                => attachment.filename,
								   :src                  => preview_office_url(attachment),
								   :onload               => "resizeObject(this);"
								  }.merge(options)
							 ),
							 :style => "position:relative;padding-top:141%;"
                )
              end #case

          end #def
                                        
        end #base
      end #self

      module InstanceMethods    
      end #module
      
    end #module
  end #module
end #module

unless ApplicationHelper.included_modules.include?(RedminePreviewOffice::Patches::ApplicationHelperPatch)
    ApplicationHelper.send(:include, RedminePreviewOffice::Patches::ApplicationHelperPatch)
end


