require_relative 'table_row_properties/table_row_height'
module OoxmlParser
  # Class for describing Table Row Properties
  class TableRowProperties < OOXMLDocumentObject
    # @return [Float] Table Row Height
    attr_accessor :height

    # Parse Columns data
    # @param [Nokogiri::XML:Element] node with Table Row Properties data
    # @return [TableRowProperties] value of Columns data
    def parse(node)
      node.xpath('*').each do |pr_child|
        case pr_child.name
        when 'trHeight'
          @height = TableRowHeight.new(parent: self).parse(pr_child)
        end
      end
      self
    end
  end
end
