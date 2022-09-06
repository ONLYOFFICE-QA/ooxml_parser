# frozen_string_literal: true

module OoxmlParser
  # method to help to work with SlideMasters
  module SlideMastersHelper
    # @return [Array<String>] list of slide masters files
    def slide_masters_files
      Dir["#{root_object.unpacked_folder}ppt/slideMasters/*.xml"]
    end

    private

    # Parse slide masters file
    def parse_slide_masters
      slide_masters_files.each do |file|
        @slide_masters << SlideMasterFile.new(parent: self).parse(file)
      end
    end
  end
end
