module OoxmlParser
  # Fallback DOCX Shape Properties
  class OldDocxShapeProperties < OOXMLDocumentObject
    attr_accessor :fill_color, :opacity, :stroke_color, :stroke_weight

    # Parse OldDocxShapeProperties object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [OldDocxShapeProperties] result of parsing
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'fillcolor'
          @fill_color = Color.new(parent: self).parse_hex_string(value.value.delete('#'))
        when 'opacity'
          @opacity = value.value.to_f
        when 'strokecolor'
          @stroke_color = Color.new(parent: self).parse_hex_string(value.value.delete('#'))
        when 'strokeweight'
          @stroke_weight = value.value.to_f
        end
      end
      self
    end
  end
end
