# frozen_string_literal: true

require_relative 'presentation_theme/font_scheme'
require_relative 'presentation_theme/theme_color'
module OoxmlParser
  # Class for data for PresentationTheme
  class PresentationTheme < OOXMLDocumentObject
    attr_accessor :name, :color_scheme
    # @return [FontScheme] font scheme
    attr_accessor :font_scheme

    def initialize(parent: nil)
      @name = ''
      @color_scheme = {}
      super
    end

    # Parse PresentationTheme
    # @param file [String] path to file to parse
    # @return [PresentationTheme] result of parsing
    def parse(file)
      root_object.add_to_xmls_stack(file)
      unless File.exist?(root_object.current_xml)
        root_object.xmls_stack.pop
        return
      end
      doc = parse_xml(root_object.current_xml)

      doc.xpath('a:theme').each do |theme_node|
        @name = theme_node.attribute('name').value if theme_node.attribute('name')
        theme_node.xpath('a:themeElements/*').each do |theme_element_node|
          case theme_element_node.name
          when 'clrScheme'
            theme_element_node.xpath('*').each do |color_scheme_element|
              @color_scheme[color_scheme_element.name.to_sym] = ThemeColor.new.parse(color_scheme_element)
            end
            @color_scheme[:background1] = @color_scheme[:lt1]
            @color_scheme[:background2] = @color_scheme[:lt2]
            @color_scheme[:bg1] = @color_scheme[:lt1]
            @color_scheme[:bg2] = @color_scheme[:lt2]
            @color_scheme[:text1] = @color_scheme[:dk1]
            @color_scheme[:text2] = @color_scheme[:dk2]
            @color_scheme[:tx1] = @color_scheme[:dk1]
            @color_scheme[:tx2] = @color_scheme[:dk2]
          when 'fontScheme'
            @font_scheme = FontScheme.new(parent: self).parse(theme_element_node)
          end
        end
      end
      root_object.xmls_stack.pop
      self
    end
  end
end
