require_relative 'ooxml_shape_body_properties/preset_text_warp'

module OoxmlParser
  # Class for parsing `bodyPr`
  class OOXMLShapeBodyProperties < OOXMLDocumentObject
    attr_accessor :margins, :anchor, :wrap, :preset_text_warp
    # @return [Symbol] Vertical Text
    attr_accessor :vertical

    alias vertical_align anchor

    # Parse OOXMLShapeBodyProperties
    # @param [Nokogiri::XML:Node] node with OOXMLShapeBodyProperties
    # @return [OOXMLShapeBodyProperties] result of parsing
    def parse(node)
      @margins = ParagraphMargins.new(OoxmlSize.new(0.127, :centimeter),
                                      OoxmlSize.new(0.127, :centimeter),
                                      OoxmlSize.new(0.254, :centimeter),
                                      OoxmlSize.new(0.254, :centimeter)).parse(node)
      node.attributes.each do |key, value|
        case key
        when 'wrap'
          @wrap = value.value.to_sym
        when 'anchor'
          @anchor = value_to_symbol(value)
        when 'vert'
          @vertical = value_to_symbol(value)
        end
      end
      node.xpath('*').each do |node_child|
        case node_child.name
        when 'prstTxWarp'
          @preset_text_warp = PresetTextWarp.parse(node_child)
        end
      end
      self
    end
  end
end
