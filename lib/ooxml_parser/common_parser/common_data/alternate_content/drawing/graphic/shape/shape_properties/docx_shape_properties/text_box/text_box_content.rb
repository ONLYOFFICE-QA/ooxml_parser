module OoxmlParser
  # Class for working with TextBoxContent (w:txbxContent)
  class TextBoxContent < OOXMLDocumentObject
    # @return [Array] list of elements
    attr_reader :elements

    def initialize(parent: nil)
      @elements = []
      @parent = parent
    end

    # Parse TextBoxContent object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [TextBoxContent] result of parsing
    def parse(node)
      node.xpath('*').each do |node_child|
        case node_child.name
        when 'p'
          @elements << Paragraph.new(parent: self).parse(node_child)
        when 'tbl'
          @elements << Table.new(parent: self).parse(node_child)
        end
      end
      self
    end
  end
end
