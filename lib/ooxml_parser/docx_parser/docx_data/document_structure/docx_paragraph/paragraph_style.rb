require_relative 'paragraph_style/bookmark'
require_relative 'paragraph_style/style_parametres'
require_relative 'paragraph_style/paragraph_tab'
# Paragraph Style Data
module OoxmlParser
  class ParagraphStyle
    attr_accessor :name,
                  :font_name,
                  :font_size,
                  :font_style,
                  :spacing,
                  :indent,
                  :borders,
                  :background

    def initialize(style_name = nil, font_name = nil, font_size = nil, font_style = nil, spacing = nil,
                   indent = nil, borders = nil, background = nil)
      @name = style_name if style_name
      @font_name = font_name if font_name
      @font_size = font_size if font_size
      @font_style = font_style if font_style
      @spacing = spacing if spacing
      @indent = indent if indent
      @borders = borders if borders
      @background = background if background
    end

    def string_style_name
      @name.to_s.tr('_', ' ')
    end

    def self.init_by_name(style_name, object = :docx_paragraph)
      found_style = nil
      AllTestData.paragraph_settings::PARAGRAPH_STYLE_DATA.each do |style|
        found_style = style.clone if style.name == style_name
      end

      fail "Cannot find this style name - #{style_name}" if found_style.nil?

      case object
      when :table
        found_style.spacing = Spacing.new(found_style.spacing.before, 0, 1, found_style.spacing.line_rule)
      when :list
        found_style.indent = Indents.new(-0.635, 1.27, 0)
      end

      found_style
    end

    def ==(other)
      # FIXME, correct border comparision
      @font_name == other.font_name && @font_size.to_f == other.font_size.to_f && @font_style == other.font_style &&
        @spacing == other.spacing && @indent == other.indent && @background == other.background # && @border == other.border
    end
  end
end
