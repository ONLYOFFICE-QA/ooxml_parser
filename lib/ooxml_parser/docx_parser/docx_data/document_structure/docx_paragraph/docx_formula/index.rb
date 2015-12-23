# Index data
module OoxmlParser
  class Index
    attr_accessor :value, :top_index, :bottom_index

    def initialize(value = nil, bottom_index = nil, top_index = nil)
      @value = value
      @bottom_index = bottom_index
      @top_index = top_index
    end
  end
end
