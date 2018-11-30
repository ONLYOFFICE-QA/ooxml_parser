module OoxmlParser
  # Class for parsing `tableStyles.xml` file
  class TableStyles < OOXMLDocumentObject
    # @return [Array<TableStyle>] list of table styles
    attr_reader :table_style_list

    def initialize(parent: nil)
      @table_style_list = []
      @parent = parent
    end

    # Parse TableStyles object
    # @param file [Nokogiri::XML:Element] node to parse
    # @return [TableStyles] result of parsing
    def parse(file = "#{OOXMLDocumentObject.path_to_folder}/#{OOXMLDocumentObject.root_subfolder}/tableStyles.xml")
      return nil unless File.exist?(file)

      document = parse_xml(file)
      node = document.xpath('*').first

      node.xpath('*').each do |node_child|
        case node_child.name
        when 'tblStyle'
          @table_style_list << TableStyle.new(parent: self).parse(node_child)
        end
      end
      self
    end

    # @param id [String] style to find
    # @return [TableStyle] style by this id
    def style_by_id(id)
      table_style_list.detect { |style| style.id == id }
    end
  end
end
