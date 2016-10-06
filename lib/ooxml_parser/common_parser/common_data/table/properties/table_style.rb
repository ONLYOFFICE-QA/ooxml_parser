require_relative 'table_element'
# Style of Table
module OoxmlParser
  class TableStyle < OOXMLDocumentObject
    attr_accessor :id, :name, :whole_table, :banding_1_horizontal, :banding_2_horizontal, :banding_1_vertical, :banding_2_vertical,
                  :last_column, :first_column, :last_row, :first_row, :southeast_cell, :southwest_cell, :northeast_cell,
                  :northwest_cell

    def initialize(id = nil)
      @id = id
    end

    def self.parse(style_id: nil, table_style_node: nil)
      table_style = TableStyle.new
      if !style_id.nil?
        nodes = get_style_node(style_id)
        return nil if nodes.nil?
        xpath = '*'
        attributes = 'style_node_child.name'
      elsif !table_style_node.nil?
        nodes = table_style_node
        xpath = 'w:tblStylePr'
        attributes = "style_node_child.attribute('type').value"
      end
      begin
        nodes.xpath(xpath).each do |style_node_child|
          case instance_eval(attributes)
          when 'wholeTbl'
            table_style.whole_table = TableElement.new(parent: table_style).parse(style_node_child)
          when 'band1H'
            table_style.banding_1_horizontal = TableElement.new(parent: table_style).parse(style_node_child)
          when 'band2H', 'band2Horz'
            table_style.banding_2_horizontal = TableElement.new(parent: table_style).parse(style_node_child)
          when 'band1V'
            table_style.banding_1_vertical = TableElement.new(parent: table_style).parse(style_node_child)
          when 'band2V'
            table_style.banding_2_vertical = TableElement.new(parent: table_style).parse(style_node_child)
          when 'lastCol'
            table_style.last_column = TableElement.new(parent: table_style).parse(style_node_child)
          when 'firstCol'
            table_style.first_column = TableElement.new(parent: table_style).parse(style_node_child)
          when 'lastRow'
            table_style.last_row = TableElement.new(parent: table_style).parse(style_node_child)
          when 'firstRow'
            table_style.first_row = TableElement.new(parent: table_style).parse(style_node_child)
          when 'seCell'
            table_style.southeast_cell = TableElement.new(parent: table_style).parse(style_node_child)
          when 'swCell'
            table_style.southwest_cell = TableElement.new(parent: table_style).parse(style_node_child)
          when 'neCell'
            table_style.northeast_cell = TableElement.new(parent: table_style).parse(style_node_child)
          when 'nwCell'
            table_style.northwest_cell = TableElement.new(parent: table_style).parse(style_node_child)
          end
        end
      end
      table_style.id = style_id unless style_id.nil?
      table_style
    end

    def self.get_style_node(style_id)
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
