module OoxmlParser
  class Columns
    attr_accessor :columns, :separator

    def initialize(columns_count)
      @columns = Array.new(columns_count)
    end
  end
end
