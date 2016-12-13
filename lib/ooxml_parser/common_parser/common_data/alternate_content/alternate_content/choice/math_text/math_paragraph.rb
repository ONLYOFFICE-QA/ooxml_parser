module OoxmlParser
  # Class for storing math paragraph
  class MathParagraph < OOXMLDocumentObject
    attr_accessor :math
    # Parse MathParagraph object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [MathParagraph] result of parsing
    def parse(node)
      node.xpath('*').each do |node_child|
        case node_child.name
        when 'oMath'
          @math = DocxFormula.new(parent: self).parse(node_child)
        end
      end
      self
    end
  end
end
