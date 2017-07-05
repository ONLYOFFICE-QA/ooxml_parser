module OoxmlParser
  # Class for parsing `w:w:sdtContent` tags
  class SDTContent < OOXMLDocumentObject
    # @return [Array, ParagraphRun] runs of sdt
    attr_reader :runs

    def initialize(parent: nil)
      @runs = []
      @parent = parent
    end

    # Parse SDTContent object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [SDTContent] result of parsing
    def parse(node)
      node.xpath('*').each do |node_child|
        case node_child.name
        when 'r'
          @runs << ParagraphRun.new(parent: self).parse(node_child)
        end
      end
      self
    end
  end
end
