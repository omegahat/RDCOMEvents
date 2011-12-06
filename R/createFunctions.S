#
# Todo:  
#   return values  - 
#   out parameters - don't change them.

# done
#   validate function  - number of arguments. (See mergeFunctions)


createFunctionTemplates =
#
#  Takes a collection of functions (typically from getFuncs(lib[[ ]]))  
#  and generates the  template definitions.

#  One can specify a different operator. This is just a simple call to lapply.

function(funcs, op = createCompiledFunction, ...)
{
  lapply(funcs, op, ...)
}	

createFunction =
#
#  Create a simple template function that just has 
#  the parameter names in the formals().
# 
#XXX Put the return value in.
#
function(func, ...)
{
  txt = paste("function(", paste(names(func@parameters), collapse = ", "), ")\n{ \n}")
  f = eval(parse(text = txt))

  class(f) = c("RDCOMProxyFunction", "RDCOMGeneratedFunction", class(f))
  attr(f, "ParameterNames") = names(func@parameters)

  returnValue = getDefaultReturnValue(func@returnType)
  if(!is.null(returnValue)) {
    attr(f, "ReturnValue") = returnValue$value
    attr(f, "DCOMReturnType") = returnValue$type
  }

  f
}

setGeneric("getDefaultReturnValue", function(type) standardGeneric("getDefaultReturnValue"))

ReturnValues = 
list("HRESULT" = list(value = as.integer(0), type = "integer"))

setMethod("getDefaultReturnValue", "TypeDescription",
           function(type) {
               idx =  match(type@name, names(ReturnValues))
	       if(is.na(idx))
                 return(NULL)

               ReturnValues[[idx]]
           })
          


createCompiledFunction =
#
# Generates an R function template from the function description (func)
# from the type library. The lib parameter gives us the type library
# in case we need to look up any value while generate the code.

function(func, lib)
{
  f = createFunction(func)

  expr = lapply(func@parameters,
                function(x) generateCoercion(x@name, x@type, lib))

  nulls = sapply(expr, is.null)
  if(any(nulls))
    expr = expr[!nulls]

  attr(f, "ParameterContracts") = expr

  f
}


# Generic function that gets the name of the R type to be used in a
# call to as(obj, <type>)
# or alternatively a method can return an expression if it needs to 
# do something slightly different, e.g. as.numeric(obj)
# The name is the name of the variable.
setGeneric("getRTypeName", function(type, name) standardGeneric("getRTypeName"))

setMethod("getRTypeName", "TypeDescriptionRef", 
             function(type, name)
               names(type@reftype)
         )

TypeDescriptionMap = 
 list("IDispatch*" = NULL,
      "VARIANT_BOOL" = "logical"
     )

setMethod("getRTypeName", "TypeDescription", 
             function(type, name) {
               which = match(type@name, names(TypeDescriptionMap))
               if(is.na(which))
                  NULL
               else
                  TypeDescriptionMap[[which]]
             })

# Out variable. Leave as is.
setMethod("getRTypeName", "PointerTypeDescription", 
             function(type, name)
              NULL
          )

generateCoercion =
#
# This generates the expression for the contract to coerce the 
# specified variable to the given DCOM type in R.
#
  #
  # We put the expression inside a quote so that it can be dput'ed correctly.
  # Otherwise, we get the expression in raw source. This is a problem of dput, but ....
  #
function(name, type, lib)
{
  type = getRTypeName(type, name)
  if(is.character(type) && length(type))
    type = substitute( quote(obj <- as(obj, type), list(obj = as.name(name),
                                                  type = type)))

  type
}

# Deal with return type


changeVariables = 
#
# Change the variable names used in the contract coercions to 
# those given in the second argument which are the parameter
# names of the function in which these contract expressions are
# to be used.
#
# Returns a list of expressions.
# We put these into the body later.
function(funcDef, varNames)
{
  # Handle ...  Discard all argument names from there on, inclusive.
  idx = match("...", varNames)
  if(!is.na(idx))
    varNames = varNames[seq(1, idx - 1)]

  if(length(names(varNames)) == 0)
    names(varNames) = names(funcDef)

  e = lapply(names(funcDef), 
                    function(name) {
                       tmp = list( as.name(varNames[name]) )
                       names(tmp) = name
                       do.call("substitute", list(funcDef[[name]], tmp))
                    })

  names(e) = names(funcDef)

  e
}


insertContract =
#
# This puts  the contract expressions into the function
# at the top of the body.
# It does this by first modifying the contract expressions
# to use the names of the parameters in the functions.
#(It assumes the same order of the parameters and the elements of the contract.)
# Then it creates a new body and puts the contract expressions there and then
# copies the expressions in the body of the original function into that.
# Then it puts this new body onto the original function and repairs the 
# environment.
function(fun, contract)
{
  contract = changeVariables(contract, names(formals(fun)))

  if(length(contract) == 0)
    return(fun)

  env = environment(fun)
  b = body(function(){})

                    # get the expression out of the quote.
  b[seq(2, length = length(contract))] = lapply(contract, function(x) x[[2]])
  
  bfun = body(fun)
  b[seq(length(b) + 1, length = length(bfun)-1)] = bfun[2:length(bfun)]


  body(fun) = b
  environment(fun) = env

  fun
}


mergeFunctions =
#
# This is the top-level function that takes the 
# user's functions and merges the information
# from the template.
#
function(defs, templates)
{
  idx = match(names(defs), names(templates))
  if(any(is.na(idx))) {
    warning("ignoring function elements ", paste(names(defs)[is.na(idx)], sep = ", "))
    defs = defs[!is.na(idx)]
  }

  funs = lapply(names(defs),
                 function(id) {
                   def = defs[[id]]
                   template = templates[[id]]

                   parms = formals(def)
                   if(length(parms) < length(formals(template)) && !("..." %in% names(parms))) 
                      stop("Too few arguments in user-supplied", id, ". Expecting ", length(formals(template)))

                   f = insertContract(defs[[id]], attr(templates[[id]], "ParameterContracts"))

		   if("ReturnValue" %in% names(attributes(template))) {
                      attr(f, "ReturnValue") = attr(template, "ReturnValue")
                      attr(f, "DCOMReturnType") = attr(template, "DCOMReturnType")
                   }
                   f
                 })
  names(funs) = names(defs)

  funs
}



