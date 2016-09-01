require_relative 'paragrpah_properties/numbering_properties'
require_relative 'paragrpah_properties/paragraph_borders'
require_relative 'paragrpah_properties/spacing'
module OoxmlParser
  class ParagraphProperties < OOXMLDocumentObject
    attr_accessor :align, :numbering, :level, :spacing, :spacing_before, :spacing_after, :indent, :margin_left, :margin_right,
                  :tabs
    # @return [RunProperties] properties of run
    attr_accessor :run_properties
    # @return [Borders] borders of paragraph
    attr_accessor :paragraph_borders
    # @return [True, False] Specifies that the paragraph
    # (or at least part of it) should be rendered on
    # the same page as the next paragraph when possible
    attr_accessor :keep_next

    def initialize(numbering = NumberingProperties.new)
      @numbering = numbering
      @spacing = Spacing.new(0, 0, 1, :multiple)
      @keep_next = false
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
          paragraph_properties.indent = OoxmlSize.new(value.value.to_f, :emu)
        when 'marR'
          paragraph_properties.margin_right = OoxmlSize.new(value.value.to_f, :emu)
        when 'marL'
          paragraph_properties.margin_left = OoxmlSize.new(value.value.to_f, :emu)
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
        when 'ind'
          paragraph_properties.indent = Indents.new(parent: paragraph_properties).parse(properties_element)
        when 'rPr'
          paragraph_properties.run_properties = RunProperties.new(parent: paragraph_properties).parse(properties_element)
        when 'pBdr'
          paragraph_properties.paragraph_borders = ParagraphBorders.parse(properties_element)
        when 'keepNext'
          paragraph_properties.keep_next = true
        end
      end
      paragraph_properties
    end
  end
end
