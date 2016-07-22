require_relative 'paragrpah_properties/numbering_properties'
require_relative 'paragrpah_properties/spacing'
module OoxmlParser
  class ParagraphProperties < OOXMLDocumentObject
    attr_accessor :align, :numbering, :level, :spacing, :spacing_before, :spacing_after, :indent, :margin_left, :margin_right,
                  :tabs

    def initialize(numbering = NumberingProperties.new)
      @numbering = numbering
      @spacing = Spacing.new(0, 0, 1, :multiple)
    end

    def self.parse(paragraph_props_node)
      paragraph_properties = ParagraphProperties.new
      paragraph_properties.spacing.parse(paragraph_props_node)
      paragraph_props_node.attributes.each do |key, value|
        case key
        when 'algn'
          paragraph_properties.align = Alignment.parse(value)
        when 'lvl'
          paragraph_properties.level = value.value.to_i
        when 'indent'
          paragraph_properties.indent = (value.value.to_f / 360_000.0).round(2)
        when 'marR'
          paragraph_properties.margin_right = (value.value.to_f / 360_000.0).round(2)
        when 'marL'
          paragraph_properties.margin_left = (value.value.to_f / 360_000.0).round(2)
        end
      end
      paragraph_props_node.xpath('*').each do |properties_element|
        case properties_element.name
        when 'buSzPct'
          paragraph_properties.numbering.size = properties_element.attribute('val').value
        when 'buFont'
          paragraph_properties.numbering.font = properties_element.attribute('typeface').value
        when 'buChar'
          paragraph_properties.numbering.symbol = properties_element.attribute('char').value
        when 'buAutoNum'
          paragraph_properties.numbering.type = properties_element.attribute('type').value.to_sym
          paragraph_properties.numbering.start_at = properties_element.attribute('startAt').value if properties_element.attribute('startAt')
        when 'tabLst'
          paragraph_properties.tabs = ParagraphTab.parse(properties_element)
        end
      end
      paragraph_properties
    end
  end
end
