# Docx Shape Body Properties
require_relative 'ooxml_shape_body_properties/preset_text_warp'

module OoxmlParser
  class OOXMLShapeBodyProperties < OOXMLDocumentObject
    attr_accessor :margins, :anchor, :wrap, :preset_text_warp

    alias vertical_align anchor

    def self.parse(body_pr_node)
      body_properties = OOXMLShapeBodyProperties.new
      body_properties.margins = ParagraphMargins.new(OoxmlSize.new(0.127, :centimeter),
                                                     OoxmlSize.new(0.127, :centimeter),
                                                     OoxmlSize.new(0.254, :centimeter),
                                                     OoxmlSize.new(0.254, :centimeter)).parse(body_pr_node)
      body_pr_node.attributes.each do |key, value|
        case key
        when 'wrap'
          body_properties.wrap = value.value.to_sym
        when 'anchor'
          body_properties.anchor = Alignment.parse(value)
        end
      end
      body_pr_node.xpath('*').each do |body_properties_child|
        case body_properties_child.name
        when 'prstTxWarp'
          body_properties.preset_text_warp = PresetTextWarp.parse(body_properties_child)
        end
      end
      body_properties
    end
  end
end
