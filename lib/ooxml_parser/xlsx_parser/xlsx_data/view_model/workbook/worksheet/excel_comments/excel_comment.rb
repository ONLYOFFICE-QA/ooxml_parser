module OoxmlParser
  # Single Comment of XLSX
  class ExcelComment < OOXMLDocumentObject
    attr_accessor :characters

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
      self
    end
  end
end
