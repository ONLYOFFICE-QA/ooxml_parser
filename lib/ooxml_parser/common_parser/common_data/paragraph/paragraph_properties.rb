require_relative 'paragrpah_properties/numbering_properties'
require_relative 'paragrpah_properties/paragraph_borders'
require_relative 'paragrpah_properties/paragraph_spacing'
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
    # @return [PageProperties] properties of section
    attr_accessor :section_properties

    def initialize(numbering = NumberingProperties.new, parent: nil)
      @numbering = numbering
      @spacing = Spacing.new(0, 0, 1, :multiple)
      @keep_next = false
      @parent = parent
    end

    # Parse ParagraphProperties object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [ParagraphProperties] result of parsing
    def parse(node)
      @spacing.parse(node)
      node.attributes.each do |key, value|
        case key
        when 'algn'
          @align = Alignment.parse(value)
        when 'lvl'
          @level = value.value.to_i
        when 'indent'
          @indent = OoxmlSize.new(value.value.to_f, :emu)
        when 'marR'
          @margin_right = OoxmlSize.new(value.value.to_f, :emu)
        when 'marL'
          @margin_left = OoxmlSize.new(value.value.to_f, :emu)
        end
      end
      node.xpath('*').each do |node_child|
        case node_child.name
        when 'buSzPct'
          @numbering.size = node_child.attribute('val').value
        when 'buFont'
          @numbering.font = node_child.attribute('typeface').value
        when 'buChar'
          @numbering.symbol = node_child.attribute('char').value
        when 'buAutoNum'
          @numbering.type = node_child.attribute('type').value.to_sym
          @numbering.start_at = node_child.attribute('startAt').value if node_child.attribute('startAt')
        when 'tabLst'
          @tabs = ParagraphTab.parse(node_child)
        when 'ind'
          @indent = Indents.new(parent: self).parse(node_child)
        when 'rPr'
          @run_properties = RunProperties.new(parent: self).parse(node_child)
        when 'pBdr'
          @paragraph_borders = ParagraphBorders.parse(node_child)
        when 'keepNext'
          @keep_next = true
        when 'sectPr'
          @section_properties = PageProperties.new(parent: self).parse(node_child, @parent, DocxParagraphRun.new)
        when 'spacing'
          @spacing = ParagraphSpacing.new(parent: self).parse(node_child)
        end
      end
      self
    end
  end
end
