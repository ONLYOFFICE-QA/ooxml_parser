module OoxmlParser
  # Class for working with TextBoxContent (w:txbxContent)
  class TextBoxContent < OOXMLDocumentObject
    # @return [Array<Paragraphs>] list of paragraphs
    attr_reader :paragraphs
    # @return [Array<Table>] list of tables
    attr_reader :tables

    def initialize(parent: nil)
      @paragraphs = []
      @tables = []
      @parent = parent
    end

    # Parse TextBoxContent object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [TextBoxContent] result of parsing
    def parse(node)
      node.xpath('*').each do |node_child|
        case node_child.name
        when 'p'
          @paragraphs << Paragraph.new(parent: self).parse(node_child)
        when 'tbl'
          @tables << Table.new(parent: self).parse(node_child)
        end
      end
      self
    end

    # @return [Array] list of all elements in text box control
    def elements
      [@paragraphs, @tables].flatten
    end
  end
end
