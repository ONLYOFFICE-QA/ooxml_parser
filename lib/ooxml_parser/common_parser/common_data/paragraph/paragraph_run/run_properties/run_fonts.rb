# frozen_string_literal: true

module OoxmlParser
  # Class for parsing `w:rFonts` object
  class RunFonts < OOXMLDocumentObject
    # @return [String] ascii font
    attr_accessor :ascii
    # @return [String] ascii theme value
    attr_accessor :ascii_theme

    # Parse RunFonts
    # @param [Nokogiri::XML:Node] node with RunFonts
    # @return [RunFonts] result of parsing
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'ascii'
          @ascii = value.to_s
        when 'asciiTheme'
          @ascii_theme = value.to_s
        end
      end
      self
    end
  end
end
