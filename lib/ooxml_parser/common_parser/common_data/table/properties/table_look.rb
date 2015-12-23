# Describe look of table
module OoxmlParser
  class TableLook < OOXMLDocumentObject
    attr_accessor :first_row, :first_column, :last_row, :last_column, :banding_row, :banding_column, :no_horizontal_banding,
                  :no_vertical_banding

    def initialize
      @first_row = false
      @first_column = false
      @last_row = false
      @last_column = false
      @banding_row = false
      @banding_column = false
      @no_horizontal_banding = false
      @no_horizontal_banding = false
    end

    def self.parse(table_look_node)
      table_look = TableLook.new
      table_look_node.attributes.each do |key, value|
        case key
        when 'firstRow'
          table_look.first_row = OOXMLDocumentObject.option_enabled?(value)
        when 'lastRow'
          table_look.last_row = OOXMLDocumentObject.option_enabled?(value)
        when 'firstCol', 'firstColumn'
          table_look.first_column = OOXMLDocumentObject.option_enabled?(value)
        when 'lastCol', 'lastColumn'
          table_look.last_column = OOXMLDocumentObject.option_enabled?(value)
        when 'noHBand'
          table_look.no_horizontal_banding = OOXMLDocumentObject.option_enabled?(value)
        when 'noVBand'
          table_look.no_vertical_banding = OOXMLDocumentObject.option_enabled?(value)
        when 'bandRow'
          table_look.banding_row = OOXMLDocumentObject.option_enabled?(value)
        when 'bandCol'
          table_look.banding_column = OOXMLDocumentObject.option_enabled?(value)
        end
      end
      table_look
    end
  end
end
