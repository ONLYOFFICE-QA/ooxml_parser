# Box data
module OoxmlParser
  class Box
    attr_accessor :borders, :element

    def initialize(borders = false, element = nil)
      @borders = borders
      @element = element
    end
  end
end
