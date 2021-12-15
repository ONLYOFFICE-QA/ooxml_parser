# frozen_string_literal: true

module OoxmlParser
  # Class for parsing <sheet> tag
  class Sheet < OOXMLDocumentObject
    # @return [String] Name of sheet
    attr_reader :name
    # @return [Integer] SheetId of sheet
    attr_reader :sheet_id
    # @return [Symbol] Specifies if sheet is hidden
    attr_reader :state
    # @return [String] Id of sheet
    attr_reader :id

    def initialize(parent: nil)
      @state = :visible
      super
    end

    # Parse Sheet data
    # @param [Nokogiri::XML:Element] node with Sheet data
    # @return [Sheet] value of Sheet
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'name'
          @name = value.value.to_s
        when 'sheetId'
          @sheet_id = value.value.to_i
        when 'state'
          @state = value.value.to_sym
        when 'id'
          @id = value.value.to_s
        end
      end
      self
    end
  end
end
