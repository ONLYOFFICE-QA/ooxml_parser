require_relative 'autofilter/filter_column'
module OoxmlParser
  # Class for `autoFilter` data
  # AutoFilter temporarily hides rows based on a
  # filter criteria, which is applied column by column
  # to a table of data in the worksheet.
  # This collection expresses AutoFilter settings.
  class Autofilter < OOXMLDocumentObject
    # @return [Coordinates] Reference to the cell range to which the AutoFilter is applied.
    attr_accessor :ref
    # @return [FilterColumn] data of filter column
    attr_accessor :filter_column

    # Parse Autofilter data
    # @param [Nokogiri::XML:Element] node with Autofilter data
    # @return [Autofilter] value of Autofilter data
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'ref'
          @ref = Coordinates.parser_coordinates_range(value.value)
        end
      end

      node.xpath('*').each do |node_child|
        case node_child.name
        when 'filterColumn'
          @filter_column = FilterColumn.new(parent: self).parse(node_child)
        end
      end
      self
    end
  end
end
