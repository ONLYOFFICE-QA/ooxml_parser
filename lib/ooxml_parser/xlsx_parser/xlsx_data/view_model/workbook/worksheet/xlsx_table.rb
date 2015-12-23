# Table in xlsx
module OoxmlParser
  class XlsxTable < OOXMLDocumentObject
    attr_accessor :name, :display_name, :reference, :autofilter, :columns

    def self.parse(table_part_node)
      table = XlsxTable.new
      link_to_table_part_xml = OOXMLDocumentObject.get_link_from_rels(table_part_node.attribute('id').value)
      doc = Nokogiri::XML(File.open(OOXMLDocumentObject.path_to_folder + link_to_table_part_xml.gsub('..', 'xl')))
      table_node = doc.xpath('xmlns:table').first
      table_node.attributes.each do |key, value|
        case key
        when 'name'
          table.name = value.value.to_s
        when 'displayName'
          table.display_name = value.value.to_s
        when 'ref'
          table.reference = Coordinates.parser_coordinates_range value.value.to_s
        end
      end
      table_node.xpath('*').each do |table_node_child|
        case table_node_child.name
        when 'autoFilter'
          table.autofilter = Coordinates.parser_coordinates_range table_node_child.attribute('ref').value
        end
      end
      table
    end
  end
end
