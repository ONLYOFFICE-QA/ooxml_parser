# frozen_string_literal: true

module OoxmlParser
  # Class for parsing `w:headerReference`,
  # `w:footerReference` object
  class HeaderFooterReference < OOXMLDocumentObject
    # @return [String] id of reference
    attr_accessor :id
    # @return [String] type of reference
    attr_accessor :type

    # Parse HeaderFooterReference
    # @param [Nokogiri::XML:Node] node with HeaderFooterReference
    # @return [HeaderFooterReference] result of parsing
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'id'
          @id = value.to_s
        when 'type'
          @type = value.to_s
        end
      end
      self
    end
  end
end
