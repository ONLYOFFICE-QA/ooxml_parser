module OoxmlParser
  class GradientStop < OOXMLDocumentObject
    attr_accessor :position, :color

    # Parse GradientStop object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [GradientStop] result of parsing
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'pos'
          @position = value.value.to_i / 1_000
        end
      end

      node.xpath('*').each do |node_child|
        @color = case node_child.name
                 when 'prstClr', 'schemeClr'
                   ValuedChild.new(:string, parent: self).parse(node_child)
                 else
                   Color.new(parent: self).parse_color(node_child)
                 end
      end
      self
    end
  end
end
