module OoxmlParser
  class CellMerge
    attr_accessor :type, :count_of_merged_cells, :value

    def initialize(type = 'horizontal', value = nil, count_of_merged_cells = 2)
      @type = type
      @count_of_merged_cells = count_of_merged_cells
      @value = value
    end

    def self.parse_merge_hash(hash)
      merge = CellMerge.new
      merge.type = hash['type']
      merge.count_of_merged_cells = hash['count_of_merged_cells']
      merge
    end
  end
end
