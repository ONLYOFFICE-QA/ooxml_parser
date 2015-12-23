module OoxmlParser
  class TableGrid
    attr_accessor :columns, :mass_of_width_cells

    def initialize(columns = [])
      @columns = columns
    end
  end
end
