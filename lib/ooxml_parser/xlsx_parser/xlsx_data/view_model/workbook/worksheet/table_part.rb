require_relative 'table_part/autofilter'
require_relative 'table_part/extension_list'
require_relative 'table_part/table_style_info'
require_relative 'table_part/table_columns'
module OoxmlParser
  # Class for `tablePart` data
  class TablePart < OOXMLDocumentObject
    attr_accessor :name, :display_name, :reference, :autofilter, :columns
    # @return [ExtensionList] list of extensions
    attr_accessor :extension_list
    # @return [TableColumns] list of table columns
    attr_reader :table_columns
    # @return [TableStyleInfo] describe style of table
    attr_accessor :table_style_info

    # Parse TablePart object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [TablePart] result of parsing
    def parse(node)
      link_to_table_part_xml = OOXMLDocumentObject.get_link_from_rels(node.attribute('id').value)
      doc = parse_xml(OOXMLDocumentObject.path_to_folder + link_to_table_part_xml.gsub('..', 'xl'))
      table_node = doc.xpath('xmlns:table').first
      table_node.attributes.each do |key, value|
        case key
        when 'name'
          @name = value.value.to_s
        when 'displayName'
          @display_name = value.value.to_s
        when 'ref'
          @reference = Coordinates.parser_coordinates_range value.value.to_s
        end
      end
      table_node.xpath('*').each do |node_child|
        case node_child.name
        when 'autoFilter'
          @autofilter = Autofilter.new(parent: self).parse(node_child)
        when 'extLst'
          @extension_list = ExtensionList.new(parent: self).parse(node_child)
        when 'tableColumns'
          @table_columns = TableColumns.new(parent: self).parse(node_child)
        when 'tableStyleInfo'
          @table_style_info = TableStyleInfo.new(parent: self).parse(node_child)
        end
      end
      self
    end
  end
end
