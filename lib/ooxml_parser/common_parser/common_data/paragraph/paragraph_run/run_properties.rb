require_relative 'run_properties/language'
require_relative 'run_properties/outline'
require_relative 'run_properties/position'
require_relative 'run_properties/run_size'
require_relative 'run_properties/run_spacing'
require_relative 'run_properties/strikeout'
module OoxmlParser
  class RunProperties < OOXMLDocumentObject
    attr_accessor :font_style, :font_color, :space, :dirty, :font_name, :font_size, :baseline, :hyperlink, :caps,
                  :vertical_align, :outline
    # @return [RunSize] get run size
    attr_accessor :size
    # @return [RunSpacing] get run spacing
    attr_accessor :spacing
    # @return [RunSpacing] get color
    attr_accessor :color
    # @return [Language] language property
    attr_accessor :language
    # @return [Position] position property
    attr_accessor :position

    def initialize(font_name = '', font_style = FontStyle.new, font_color = nil, space = nil, baseline = :baseline)
      @font_name = font_name
      @font_style = font_style
      @font_color = font_color
      @space = space
      @baseline = baseline
    end

    def self.parse(character_props_node)
      character_properties = RunProperties.new
      character_properties.font_style = Presentation.current_font_style.dup
      character_props_node.attributes.each do |key, value|
        case key
        when 'sz'
          character_properties.font_size = value.value.to_f / 100.0
        when 'spc'
          character_properties.space = (value.value.to_f / 2_834.0).round(2)
        when 'b'
          character_properties.font_style.bold = OOXMLDocumentObject.option_enabled?(character_props_node, 'b')
        when 'i'
          character_properties.font_style.italic = OOXMLDocumentObject.option_enabled?(character_props_node, 'i')
        when 'u'
          character_properties.font_style.underlined = Underline.parse(value.value)
        when 'strike'
          character_properties.font_style.strike = Strikeout.parse(value)
        when 'baseline'
          case value.value.to_i
          when -25_000, -30_000
            character_properties.baseline = :subscript
          when 30_000
            character_properties.baseline = :superscript
          when 0
            character_properties.baseline = :baseline
          end
        when 'cap'
          character_properties.caps = value.value.to_sym
        end
      end
      character_properties.font_color = DocxColorScheme.parse(character_props_node)
      character_props_node.xpath('*').each do |properties_element|
        case properties_element.name
        when 'sz'
          character_properties.size = RunSize.parse(properties_element)
        when 'spacing'
          character_properties.spacing = RunSpacing.parse(properties_element)
        when 'color'
          character_properties.color = Color.parse_color_tag(properties_element)
        when 'solidFill'
          character_properties.font_color = Color.parse_color(properties_element.xpath('*').first)
        when 'latin'
          character_properties.font_name = properties_element.attribute('typeface').value
        when 'b'
          character_properties.font_style.bold = OOXMLDocumentObject.option_enabled?(properties_element)
        when 'i'
          character_properties.font_style.italic = OOXMLDocumentObject.option_enabled?(properties_element, 'i')
        when 'u'
          character_properties.font_style.underlined = Underline.new(:single)
        when 'vertAlign'
          character_properties.vertical_align = properties_element.attribute('val').value.to_sym
        when 'rFont'
          character_properties.font_name = properties_element.attribute('val').value
        when 'rFonts'
          character_properties.font_name = properties_element.attribute('ascii').value if properties_element.attribute('ascii')
        when 'color'
          character_properties.font_color = Color.parse_color_tag(properties_element)
        when 'strike'
          character_properties.font_style.strike = OOXMLDocumentObject.option_enabled?(properties_element)
        when 'hlinkClick'
          character_properties.hyperlink = Hyperlink.parse(properties_element)
        when 'ln'
          character_properties.outline = Outline.parse(properties_element)
        when 'lang'
          character_properties.language = Language.parse(properties_element)
        when 'position'
          character_properties.position = Position.parse(properties_element)
        end
      end
      character_properties.font_color = DocxColorScheme.parse(character_props_node)
      character_properties.font_name = Presentation.default_font_typeface if character_properties.font_name.empty?
      character_properties.font_size = Presentation.default_font_size if character_properties.font_size.nil?
      character_properties
    end
  end
end
