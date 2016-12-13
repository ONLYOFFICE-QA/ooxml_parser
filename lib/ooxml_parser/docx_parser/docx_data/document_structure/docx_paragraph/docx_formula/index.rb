module OoxmlParser
  # Class for 'sSubSup', 'sSup', 'sSub' data
  class Index < OOXMLDocumentObject
    attr_accessor :value, :top_index, :bottom_index

    # Parse Index object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [Index] result of parsing
    def parse(node)
      node.xpath('*').each do |node_child|
        case node_child.name
        when 'sup'
          @top_index = DocxFormula.new(parent: self).parse(node_child)
        when 'sub'
          @bottom_index = DocxFormula.new(parent: self).parse(node_child)
        end
      end
      @value = DocxFormula.new(parent: self).parse(node)
      self
    end
  end
end
