# frozen_string_literal: true

module OoxmlParser
  # Class for parsing SlideMaster files
  class SlideMaster < OOXMLDocumentObject
    # @return [CommonSlideData] common slide data
    attr_reader :common_slide_data

    # Parse SlideMaster
    # @param file [String] path to file to parse
    # @return [SlideMaster] result of parsing
    def parse(file)
      doc = parse_xml(file)
      doc.xpath('p:sldMaster/*').each do |node_child|
        case node_child.name
        when 'cSld'
          @common_slide_data = CommonSlideData.new(parent: self).parse(node_child)
        end
      end
      self
    end
  end
end
