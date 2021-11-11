# frozen_string_literal: true

module OoxmlParser
  # method to help to work with SlideLayouts
  module SlideLayoutsHelper
    # @return [Array<String>] list of slide layouts files
    def slide_layouts_files
      Dir["#{OOXMLDocumentObject.path_to_folder}ppt/slideLayouts/*.xml"]
    end

    private

    # Parse slide layouts file
    def parse_slide_layouts
      slide_layouts_files.each do |file|
        @slide_layouts << SlideLayoutFile.new(parent: self).parse(file)
      end
    end
  end
end
