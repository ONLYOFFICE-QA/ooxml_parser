require_relative 'run_properties/language'
require_relative 'run_properties/outline'
require_relative 'run_properties/position'
require_relative 'run_properties/size'
require_relative 'run_properties/run_spacing'
require_relative 'run_properties/shade'
require_relative 'run_properties/strikeout'
module OoxmlParser
  # Data about `rPr` object
  class RunProperties < OOXMLDocumentObject
    attr_accessor :font_style, :font_color, :space, :dirty, :font_name, :font_size, :baseline, :hyperlink, :caps,
                  :vertical_align, :outline
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

    def initialize(font_name = '',
                   font_style = FontStyle.new,
                   font_color = nil,
                   space = nil,
                   baseline = :baseline,
                   parent: nil)
      @font_name = font_name
      @font_style = font_style
      @font_color = font_color
      @space = space
      @baseline = baseline
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
        when 'spc'
          @space = (value.value.to_f / 2_834.0).round(2)
        when 'b'
          @font_style.bold = option_enabled?(node, 'b')
        when 'i'
          @font_style.italic = option_enabled?(node, 'i')
        when 'u'
          @font_style.underlined = Underline.parse(value.value)
        when 'strike'
          @font_style.strike = Strikeout.parse(value)
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
      @font_color = DocxColorScheme.parse(node)
      node.xpath('*').each do |properties_element|
        case properties_element.name
        when 'sz'
          @size = Size.new.parse(properties_element)
        when 'spacing'
          @spacing = RunSpacing.parse(properties_element)
        when 'color'
          @color = Color.parse_color_tag(properties_element)
        when 'solidFill'
          @font_color = Color.parse_color(properties_element.xpath('*').first)
        when 'latin'
          @font_name = properties_element.attribute('typeface').value
        when 'b'
          @font_style.bold = option_enabled?(properties_element)
        when 'i'
          @font_style.italic = option_enabled?(properties_element, 'i')
        when 'u'
          @font_style.underlined = Underline.new(:single)
        when 'vertAlign'
          @vertical_align = properties_element.attribute('val').value.to_sym
        when 'rFont'
          @font_name = properties_element.attribute('val').value
        when 'rFonts'
          @font_name = properties_element.attribute('ascii').value if properties_element.attribute('ascii')
        when 'color'
          @font_color = Color.parse_color_tag(properties_element)
        when 'strike'
          @font_style.strike = option_enabled?(properties_element)
        when 'hlinkClick'
          @hyperlink = Hyperlink.parse(properties_element)
        when 'ln'
          @outline = Outline.parse(properties_element)
        when 'lang'
          @language = Language.parse(properties_element)
        when 'position'
          @position = Position.parse(properties_element)
        when 'shd'
          @shade = Shade.parse(properties_element)
        end
      end
      @font_color = DocxColorScheme.parse(node)
      @font_name = Presentation.default_font_typeface if @font_name.empty?
      @font_size = Presentation.default_font_size if @font_size.nil?
      self
    end
  end
end
