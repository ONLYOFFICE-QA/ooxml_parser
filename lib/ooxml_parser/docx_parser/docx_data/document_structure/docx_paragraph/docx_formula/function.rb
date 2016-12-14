module OoxmlParser
  # Class for `func` data
  class Function < OOXMLDocumentObject
    attr_accessor :name, :argument

    # Parse Function object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [Function] result of parsing
    def parse(node)
      node.xpath('*').each do |node_child|
        case node_child.name
        when 'fName'
          @name = DocxFormula.new(parent: self).parse(node_child)
        end
      end
      @argument = DocxFormula.new(parent: self).parse(node)
      self
    end
  end
end
