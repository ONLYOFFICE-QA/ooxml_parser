module OoxmlParser
  # Single Comment of XLSX
  class ExcelComment < OOXMLDocumentObject
    attr_accessor :characters
    # @return [String] text of comment
    attr_reader :text

    def initialize(parent: nil)
      @characters = []
      @parent = parent
    end

    # Parse ExcelComment object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [ExcelComment] result of parsing
    def parse(node)
      node.xpath('xmlns:text/xmlns:r').each do |node_child|
        @characters << ParagraphRun.new(parent: self).parse(node_child)
      end
      node.xpath('xmlns:text/*').each do |node_child|
        case node_child.name
        when 't'
          @text = Text.new(parent: self).parse(node_child)
        end
      end
      self
    end
  end
end
