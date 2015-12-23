# Limit Data
module OoxmlParser
  class Limit
    attr_accessor :type, :element, :limit

    def initialize(type = nil, element = nil, limit = nil)
      @type = type
      @element = element
      @limit = limit
    end
  end
end
