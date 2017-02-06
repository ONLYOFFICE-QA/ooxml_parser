require_relative 'run_properties/language'
require_relative 'run_properties/outline'
require_relative 'run_properties/position'
require_relative 'run_properties/size'
require_relative 'run_properties/run_spacing'
require_relative 'run_properties/run_style'
require_relative 'run_properties/shade'
module OoxmlParser
  # Data about `rPr` object
  class RunProperties < OOXMLDocumentObject
    attr_accessor :font_style, :font_color, :space, :dirty, :font_name, :font_size, :baseline, :hyperlink, :caps,
                  :vertical_align, :outline
    attr_accessor :shadow
    # @return [Size] get run size
    attr_accessor :size
    # @return [RunSpacing] get run spacing
    attr_accessor :spacing
    # @return [RunSpacing] get color
    attr_accessor :color
    # @return [Language] language property
    attr_accessor :language
    # @return [Position] position property
    attr_accessor :position
    # @return [Shade] shade property
    attr_accessor :shade
    # @return [RunStyle] run style
    attr_accessor :run_style
    # @return [Float]
    # This element specifies the font size which shall be applied to all
    # complex script characters in the contents of this run when displayed
    attr_accessor :font_size_complex

    def initialize(parent: nil)
      @font_name = ''
      @font_style = FontStyle.new
      @baseline = :baseline
      @parent = parent
    end

    # Parse RunProperties object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [RunProperties] result of parsing
    def parse(node)
      @font_style = Presentation.current_font_style.dup
      node.attributes.each do |key, value|
        case key
        when 'sz'
          @font_size = value.value.to_f / 100.0
        when 'szCs'
          @font_size_complex = node.attribute('val').value.to_i / 2.0
        when 'shadow'
          @shadow = true
        when 'spc'
          @space = OoxmlSize.new(value.value.to_f, :one_100th_point)
        when 'b'
          @font_style.bold = option_enabled?(node, 'b')
        when 'i'
          @font_style.italic = option_enabled?(node, 'i')
        when 'u'
          @font_style.underlined = Underline.new(parent: self).parse(value.value)
        when 'strike'
          @font_style.strike = value_to_symbol(value)
        when 'baseline'
          case value.value.to_i
          when -25_000, -30_000
            @baseline = :subscript
          when 30_000
            @baseline = :superscript
          when 0
            @baseline = :baseline
          end
        when 'cap'
          @caps = value.value.to_sym
        end
      end
      node.xpath('*').each do |node_child|
        case node_child.name
        when 'sz'
          @size = Size.new.parse(node_child)
        when 'spacing'
          @spacing = RunSpacing.new(parent: self).parse(node_child)
        when 'color'
          @color = Color.parse_color_tag(node_child)
        when 'solidFill'
          @font_color = Color.parse_color(node_child.xpath('*').first)
        when 'latin'
          @font_name = node_child.attribute('typeface').value
        when 'b'
          @font_style.bold = option_enabled?(node_child)
        when 'i'
          @font_style.italic = option_enabled?(node_child, 'i')
        when 'u'
          @font_style.underlined = Underline.new(:single)
        when 'vertAlign'
          @vertical_align = node_child.attribute('val').value.to_sym
        when 'rFont'
          @font_name = node_child.attribute('val').value
        when 'rFonts'
          @font_name = node_child.attribute('ascii').value if node_child.attribute('ascii')
        when 'strike'
          @font_style.strike = option_enabled?(node_child)
        when 'hlinkClick'
          @hyperlink = Hyperlink.new(parent: self).parse(node_child)
        when 'ln'
          @outline = Outline.new(parent: self).parse(node_child)
        when 'lang'
          @language = Language.new(parent: self).parse(node_child)
        when 'position'
          @position = Position.new(parent: self).parse(node_child)
        when 'shd'
          @shade = Shade.new(parent: self).parse(node_child)
        when 'rStyle'
          @run_style = RunStyle.new(parent: self).parse(node_child)
        end
      end
      @font_color = DocxColorScheme.new(parent: self).parse(node)
      @font_name = Presentation.default_font_typeface if @font_name.empty?
      @font_size = Presentation.default_font_size if @font_size.nil?
      self
    end
  end
end
