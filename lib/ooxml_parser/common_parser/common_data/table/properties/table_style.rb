require_relative 'table_element'
# Style of Table
module OoxmlParser
  class TableStyle < OOXMLDocumentObject
    attr_accessor :id, :name, :whole_table, :banding_1_horizontal, :banding_2_horizontal, :banding_1_vertical, :banding_2_vertical,
                  :last_column, :first_column, :last_row, :first_row, :southeast_cell, :southwest_cell, :northeast_cell,
                  :northwest_cell

    # Parse TableStyle object
    # @param style_id [String] id to parse
    # @return [TableStyle] result of parsing
    def parse(style_id: nil)
      nodes = get_style_node(style_id)
      return nil unless nodes
      nodes.xpath('*').each do |style_node_child|
        case style_node_child.name
        when 'wholeTbl'
          @whole_table = TableElement.new(parent: self).parse(style_node_child)
        when 'band1H'
          @banding_1_horizontal = TableElement.new(parent: self).parse(style_node_child)
        when 'band2H', 'band2Horz'
          @banding_2_horizontal = TableElement.new(parent: self).parse(style_node_child)
        when 'band1V'
          @banding_1_vertical = TableElement.new(parent: self).parse(style_node_child)
        when 'band2V'
          @banding_2_vertical = TableElement.new(parent: self).parse(style_node_child)
        when 'lastCol'
          @last_column = TableElement.new(parent: self).parse(style_node_child)
        when 'firstCol'
          @first_column = TableElement.new(parent: self).parse(style_node_child)
        when 'lastRow'
          @last_row = TableElement.new(parent: self).parse(style_node_child)
        when 'firstRow'
          @first_row = TableElement.new(parent: self).parse(style_node_child)
        when 'seCell'
          @southeast_cell = TableElement.new(parent: self).parse(style_node_child)
        when 'swCell'
          @southwest_cell = TableElement.new(parent: self).parse(style_node_child)
        when 'neCell'
          @northeast_cell = TableElement.new(parent: self).parse(style_node_child)
        when 'nwCell'
          @northwest_cell = TableElement.new(parent: self).parse(style_node_child)
        end
      end
      @id = style_id
      self
    end

    def get_style_node(style_id)
      begin
        doc = Nokogiri::XML(File.open(OOXMLDocumentObject.path_to_folder + 'ppt/tableStyles.xml'))
      rescue StandardError
        raise 'Can\'t find tableStyles.xml in ' + OOXMLDocumentObject.path_to_folder.to_s + '/ppt folder'
      end
      doc.xpath('a:tblStyleLst/a:tblStyle').each { |table_style_node| return table_style_node if table_style_node.attribute('styleId').value == style_id }
      nil
    end
  end
end
