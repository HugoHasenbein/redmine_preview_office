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
    module ThumbnailPatch
      def self.included(base)
        base.extend(ClassMethods)
        base.send(:include, InstanceMethods)
        base.class_eval do

          unloadable 	
          
          # for those, who read and analyze code: I haven't figured it out yet how to unset 
          # a constant and how to patch a function, which has been defined as self.function()
          # in a base.class_eval block
          #
		  @REDMINE_PREVIEW_OFFICE_CONVERT_BIN = ('libreoffice').freeze
		  
		  # Generates a thumbnail for the source image to target
		  def self.generate_preview_office(source, target, size)

			unless File.exists?(target)

			  directory = File.dirname(target)
			  unless File.exists?(directory)
				FileUtils.mkdir_p directory
			  end
			  
			  thumbnail_directory = File.join(Rails.root, "tmp", "thumbnails")
			  media_directory = File.basename(target, File.extname(target))
			                
			  Dir.mktmpdir do |tmpdir|
				cmd = "cd #{tmpdir}; #{shell_quote @REDMINE_PREVIEW_OFFICE_CONVERT_BIN} --convert-to pdf #{shell_quote source}; mv #{shell_quote File.basename(source, File.extname(source)) + ".pdf"} #{shell_quote target}"

				unless system(cmd)
				  logger.error("Creating preview with libreoffice failed (#{$?}):\nCommand: #{cmd}")
				  return nil
				end
			  end
			end
			target
		  end #def 
		                     
		  def self.libreoffice_available?
			return @libreoffice_available if defined?(@libreoffice_available)
			@libreoffice_available = system("#{shell_quote @REDMINE_PREVIEW_OFFICE_CONVERT_BIN} --version") rescue false
			logger.warn("Libre Office (#{@REDMINE_PREVIEW_OFFICE_CONVERT_BIN}) not available") unless @libreoffice_available
			@libreoffice_available
		  end

        end #base
      end #self

      module InstanceMethods          		  
      end #module
      
      module ClassMethods
      end #module
      
    end #module
  end #module
end #module

unless Redmine::Thumbnail.included_modules.include?(RedminePreviewOffice::Patches::ThumbnailPatch)
    Redmine::Thumbnail.send(:include, RedminePreviewOffice::Patches::ThumbnailPatch)
end


