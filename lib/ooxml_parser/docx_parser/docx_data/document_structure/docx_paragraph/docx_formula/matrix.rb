require_relative 'matrix/matrix_row'
module OoxmlParser
  class Matrix
    attr_accessor :rows

    def initialize(rows = [])
      @rows = rows
    end
  end
end
