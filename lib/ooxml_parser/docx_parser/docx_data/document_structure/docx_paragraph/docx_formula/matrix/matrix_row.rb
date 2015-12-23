# Matrix Row data
module OoxmlParser
  class MatrixRow
    attr_accessor :columns

    def initialize(columns_count = 1)
      @columns = Array.new(columns_count)
    end
  end
end
