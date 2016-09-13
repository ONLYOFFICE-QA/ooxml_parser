require_relative 'paragrpah_properties/numbering_properties'
require_relative 'paragrpah_properties/paragraph_borders'
require_relative 'paragrpah_properties/paragraph_spacing'
require_relative 'paragrpah_properties/spacing'
require_relative 'paragrpah_properties/tab_list'
require_relative 'paragrpah_properties/tabs'
module OoxmlParser
  class ParagraphProperties < OOXMLDocumentObject
    attr_accessor :align, :numbering, :level, :spacing, :spacing_before, :spacing_after, :indent, :margin_left, :margin_right
    # @return [Tabs] list of tabs
    attr_accessor :tabs
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
    # @return [True, False] This element specifies that any space specified
    # before or after this paragraph, specified using the spacing element (17.3.1.33),
    # should not be applied when the preceding and following paragraphs are of
    # the same paragraph style, affecting the top and bottom spacing respectively
    attr_accessor :contextual_spacing

    def initialize(numbering = NumberingProperties.new, parent: nil)
      @numbering = numbering
      @spacing = Spacing.new(OoxmlSize.new(0),
                             OoxmlSize.new(0),
                             OoxmlSize.new(1, :centimeter),
                             :multiple)
      @keep_next = false
      @tabs = []
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
          @align = value_to_symbol(value)
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
        when 'tabs'
          @tabs = Tabs.new(parent: self).parse(node_child)
        when 'tabLst'
          @tabs = TabList.new(parent: self).parse(node_child)
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
        when 'contextualSpacing'
          @contextual_spacing = option_enabled?(node_child)
        end
      end
      self
    end
  end
end
