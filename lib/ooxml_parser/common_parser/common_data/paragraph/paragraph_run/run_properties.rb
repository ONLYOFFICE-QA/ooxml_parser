# frozen_string_literal: true

require_relative 'run_properties/outline'
require_relative 'run_properties/position'
require_relative 'run_properties/run_fonts'
require_relative 'run_properties/run_spacing'
require_relative 'run_properties/run_style'
require_relative 'run_properties/shade'
require_relative 'run_properties/size'

module OoxmlParser
  # Data about `rPr` object
  class RunProperties < OOXMLDocumentObject
    # @return [FontStyle] font style of run
    attr_accessor :font_style
    # @return [Color, DocxColorScheme] color of run
    attr_reader :font_color
    # @return [OoxmlSize] space size
    attr_reader :space
    # @return [Symbol] baseline of run
    attr_reader :baseline
    # @return [Hyperlink] hyperlink of run
    attr_reader :hyperlink
    # @return [Symbol] caps data
    attr_reader :caps
    # @return [Symbol] vertical align data
    attr_reader :vertical_align
    # @return [Outline] outline data
    attr_reader :outline
    # @return [True, False] is text shadow
    attr_reader :shadow
    # @return [True, False] is text emboss
    attr_reader :emboss
    # @return [True, False] is text vanish
    attr_reader :vanish
    # @return [True, False] is text rtl
    attr_reader :rtl
    # @return [Size] get run size
    attr_accessor :size
    # @return [RunSpacing] get run spacing
    attr_accessor :spacing
    # @return [RunSpacing] get color
    attr_accessor :color
    # @return [ValuedChild] language property
    attr_reader :language
    # @return [Position] position property
    attr_accessor :position
    # @return [Shade] shade property
    attr_accessor :shade
    # @return [RunStyle] run style
    attr_accessor :run_style
    # @return [ValuedChild] ligatures type
    attr_reader :ligatures
    # @return [RunFonts] value of RunFonts
    attr_reader :run_fonts

    def initialize(params = {})
      @font_name = params.fetch(:font_name, '')
      @font_style = FontStyle.new
      @baseline = :baseline
      super(parent: params[:parent])
    end

    # Parse RunProperties object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [RunProperties] result of parsing
    def parse(node)
      @font_style = root_object.default_font_style.dup
      node.attributes.each do |key, value|
        case key
        when 'sz'
          @font_size = value.value.to_f / 100.0
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
          @baseline = parse_baseline(value)
        when 'cap'
          @caps = value.value.to_sym
        end
      end
      node.xpath('*').each do |node_child|
        case node_child.name
        when 'sz'
          @size = Size.new.parse(node_child)
        when 'shadow'
          @shadow = option_enabled?(node_child)
        when 'emboss'
          @emboss = option_enabled?(node_child)
        when 'vanish'
          @vanish = option_enabled?(node_child)
        when 'rtl'
          @rtl = option_enabled?(node_child)
        when 'spacing'
          @spacing = RunSpacing.new(parent: self).parse(node_child)
        when 'color'
          @color = OoxmlColor.new(parent: self).parse(node_child)
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
          @run_fonts = RunFonts.new(parent: self).parse(node_child)
        when 'strike'
          @font_style.strike = option_enabled?(node_child)
        when 'hlinkClick'
          @hyperlink = Hyperlink.new(parent: self).parse(node_child)
        when 'ln'
          @outline = Outline.new(parent: self).parse(node_child)
        when 'lang'
          @language = ValuedChild.new(:string, parent: self).parse(node_child)
        when 'position'
          @position = Position.new(parent: self).parse(node_child)
        when 'shd'
          @shade = Shade.new(parent: self).parse(node_child)
        when 'rStyle'
          @run_style = RunStyle.new(parent: self).parse(node_child)
        when 'ligatures'
          @ligatures = ValuedChild.new(:symbol, parent: self).parse(node_child)
        end
      end
      @font_color = DocxColorScheme.new(parent: self).parse(node)
      self
    end

    # @return [Float] font size
    def font_size
      @font_size ||= root_object.default_font_size
    end

    # @return [String] name of font
    def font_name
      return @run_fonts.ascii unless @run_fonts.ascii.empty?

      root_object.default_font_typeface
    end

    private

    # @param value [Nokogiri::XML::Attr] nokogiri parameter to parse
    # @return [Symbol] baseline value depending of type
    def parse_baseline(value)
      case value.value.to_i
      when -25_000, -30_000
        :subscript
      when 30_000
        :superscript
      when 0
        :baseline
      end
    end
  end
end
