# frozen_string_literal: true

module OoxmlParser
  # Class for parsing <clientData> tag
  class ClientData < OOXMLDocumentObject
    # @return [True, False] Specifies if drawing is locked when sheet is protected
    attr_reader :locks_with_sheet

    def initialize(parent: nil)
      @locks_with_sheet = true
      super
    end

    # Parse ClientData data
    # @param [Nokogiri::XML:Element] node with ClientData data
    # @return [Sheet] value of ClientData
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'fLocksWithSheet'
          @locks_with_sheet = attribute_enabled?(value)
        end
      end
      self
    end
  end
end
