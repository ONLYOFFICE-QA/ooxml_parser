module OoxmlParser
  class CellMerge
    attr_accessor :type, :count_of_merged_cells, :value

    def initialize(type = 'horizontal', value = nil, count_of_merged_cells = 2)
      @type = type
      @count_of_merged_cells = count_of_merged_cells
      @value = value
    end

    # Parse Merge data
    # @param [Nokogiri::XML:Element] node with Merge data
    # @return [CellMerge] value of CellMerge
    def self.parse(node)
      merge = CellMerge.new
      merge.count_of_merged_cells = node.attribute('count_rows_in_span').nil? ? nil : node.attribute('count_rows_in_span').value
      merge.value = node.attribute('val').nil? ? nil : node.attribute('val').value.to_sym
      merge.type = node.attribute('type').nil? ? nil : node.attribute('type').value.to_sym
      merge
    end
  end
end
