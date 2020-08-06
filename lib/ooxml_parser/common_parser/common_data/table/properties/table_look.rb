# frozen_string_literal: true

module OoxmlParser
  # Describe look of table, parsed `w:tblLook`
  class TableLook < OOXMLDocumentObject
    attr_accessor :first_row, :first_column, :last_row, :last_column, :banding_row, :banding_column, :no_horizontal_banding,
                  :no_vertical_banding

    def initialize(parent: nil)
      @first_row = false
      @first_column = false
      @last_row = false
      @last_column = false
      @banding_row = false
      @banding_column = false
      @no_horizontal_banding = false
      @no_horizontal_banding = false
      super
    end

    # Parse TableLook object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [TableLook] result of parsing
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'firstRow'
          @first_row = attribute_enabled?(value)
        when 'lastRow'
          @last_row = attribute_enabled?(value)
        when 'firstCol', 'firstColumn'
          @first_column = attribute_enabled?(value)
        when 'lastCol', 'lastColumn'
          @last_column = attribute_enabled?(value)
        when 'noHBand'
          @no_horizontal_banding = attribute_enabled?(value)
        when 'noVBand'
          @no_vertical_banding = attribute_enabled?(value)
        when 'bandRow'
          @banding_row = attribute_enabled?(value)
        when 'bandCol'
          @banding_column = attribute_enabled?(value)
        end
      end
      self
    end
  end
end
