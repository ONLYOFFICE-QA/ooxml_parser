# frozen_string_literal: true

module OoxmlParser
  # Class for parsing <definedName> tag
  class DefinedName < OOXMLDocumentObject
    # @return [String] Ranges to which defined name refers
    attr_reader :range
    # @return [String] Name
    attr_accessor :name
    # @return [String] Id of sheet
    attr_accessor :local_sheet_id
    # @return [Symbol] Specifies whether defined name is hidden
    attr_reader :hidden

    # Parse Defined Name data
    # @param [Nokogiri::XML:Element] node with DefinedName data
    # @return [DefinedName] value of DefinedName
    def parse(node)
      @range = node.text
      node.attributes.each do |key, value|
        case key
        when 'name'
          @name = value.value.to_s
        when 'localSheetId'
          @local_sheet_id = value.value.to_i
        when 'hidden'
          @hidden = attribute_enabled?(value)
        end
      end
      self
    end
  end
end
