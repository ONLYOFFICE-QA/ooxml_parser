# frozen_string_literal: true

module OoxmlParser
  # Class for 'm:mr' data
  class MatrixRow < OOXMLDocumentObject
    attr_accessor :columns

    def initialize(columns_count = 1, parent: nil)
      @columns = Array.new(columns_count)
      @parent = parent
    end
  end
end
