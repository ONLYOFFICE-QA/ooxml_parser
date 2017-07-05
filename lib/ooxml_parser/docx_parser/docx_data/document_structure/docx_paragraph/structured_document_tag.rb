require_relative 'sdt/sdt_content'
module OoxmlParser
  # Class for parsing `w:std`
  class StructuredDocumentTag < OOXMLDocumentObject
    # @return [SDTContent] content of sdt
    attr_reader :sdt_content

    # Parse StructuredDocumentTag object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [StructuredDocumentTag] result of parsing
    def parse(node)
      node.xpath('*').each do |node_child|
        case node_child.name
        when 'sdtContent'
          @sdt_content = SDTContent.new(parent: self).parse(node_child)
        end
      end
      self
    end
  end
end
