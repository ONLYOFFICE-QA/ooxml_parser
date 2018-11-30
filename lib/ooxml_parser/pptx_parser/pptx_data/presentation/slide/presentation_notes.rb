module OoxmlParser
  # Class for p:notes
  class PresentationNotes < OOXMLDocumentObject
    # @return [CommonSlideData] common slide data
    attr_reader :common_slide_data

    # Parse PresentationNotes object
    # @param file [String] file to parse
    # @return [PresentationNotes] result of parsing
    def parse(file)
      node = parse_xml(file)
      node.xpath('p:notes/*').each do |node_child|
        case node_child.name
        when 'cSld'
          @common_slide_data = CommonSlideData.new(parent: self).parse(node_child)
        end
      end
      self
    end
  end
end
