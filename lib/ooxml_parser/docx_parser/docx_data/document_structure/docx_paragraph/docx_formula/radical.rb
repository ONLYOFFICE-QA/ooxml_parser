module OoxmlParser
  # Class for `rad` data
  class Radical < OOXMLDocumentObject
    attr_accessor :degree, :value

    def initialize(parent: nil)
      @degree = 2
      @parent = parent
    end

    # Parse Function object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [Function] result of parsing
    def parse(node)
      node.xpath('*').each do |node_child|
        case node_child.name
        when 'deg'
          @degree = DocxFormula.new(parent: self).parse(node_child)
        end
      end
      @value = DocxFormula.new(parent: self).parse(node)
      self
    end
  end
end
