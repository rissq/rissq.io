#' Import information
#' @name import
#' @export
import <- function(metapath, datapath) {

  if(missing(metapath) || missing(datapath)) {
    stop("[Import data: ] Path to file must be specified")
  }

  characteristic <- importCharacteristic(metapath)
  process <- importProcess(metapath, list(characteristic))
  prodata <- importProData(datapath, pro = process, characteristic = characteristic)

  return(list(characteristic, process, prodata))
}

#' Import Characteristic information
#' @name importCharacteristic
#' @export
importCharacteristic <- function(metapath) {

  if(missing(metapath)) {
    stop("[Import Characteristic validation: ] Path to file must be specified")
  }

  metadata <- readxl::read_excel(metapath,
                                 sheet = "characteristic")


  Object <- .Characteristic(
    id = metadata[[1,2]],
    name = metadata[[2,2]],
    description = metadata[[3,2]],
    kvalue = metadata[[4,2]],
    T = as.numeric(metadata[[5,2]]),
    U = as.numeric(metadata[[6,2]]),
    L = as.numeric(metadata[[7,2]]),
    pnc = as.numeric(metadata[[8,2]]),
    digits = as.numeric(metadata[[9,2]]),
    units = as.numeric(metadata[[10,2]])
  )

  return(Object)
}

#' Import Process information
#' @name importProcess
#' @export
importProcess <- function(metapath, characteristics = list()) {

  if(missing(metapath)) {
    stop("[Import Process validation: ] Path to file must be specified")
  }

  metadata <- readxl::read_excel(metapath,
                                     sheet = "process")

  Object <- .Process(
    id = metadata[[1,2]],
    name = metadata[[2,2]],
    description = metadata[[3,2]],
    characteristics = characteristics
  )

  return(Object)
}

#' Import information
#' @name importProData
#' @export
importProData <- function(datapath, metadatapath = NA, pro = NA, characteristic = NA) {

  if(missing(datapath)) {
    stop("[Import ProData validation: ] Path to file must be specified")
  }

  data <-read.csv(datapath)

  if(!is.na(metadatapath) && is.na(pro)) {
    pro = importProcess(metadatapath)
  }

  if(!is.na(metadatapath) && is.na(characteristic)) {
    characteristic = importCharacteristic(metadatapath)
  }

  if(missing(pro)) {
    stop("[Import ProData validation: ] Process object or path to metadata must be given.")
  }

  if(missing(characteristic)) {
    stop("[Import ProData validation: ] Characteristic object or path to metadata must be given.")
  }

  Object <- .ProData(
    pro = pro,
    characteristic = characteristic,
    data = data
  )

  return(Object)
}

#' Import MSA information
#' @name importMSAnalysis
#' @export
importMSAnalysis <- function(metadatapath, datapath = NA, pro = NA, characteristic = NA, prodata = NA) {

  if(missing(metadatapath)) {
    stop("[Import MSA validation: ] Path to file must be specified")
  }

  metadata <- readxl::read_excel(metadatapath,
                                     sheet = "analysis")

  method <- tolower(metadata[[8,2]])

  if(!is.na(metadatapath) && is.na(pro)) {
    pro = importProcess(metadatapath)
  }

  if(!is.na(metadatapath) && is.na(characteristic)) {
    characteristic = importCharacteristic(metadatapath)
  }

  if( !is.na(metadatapath) && !is.na(datapath) && is.na(prodata)) {
    prodata = importProData(data = datapath, metadatapath = metadatapath, pro, characteristic)
  }

  if(missing(pro)) {
    stop("[Import MSA validation: ] Process object or path to metadata must be given.")
  }

  if(missing(characteristic)) {
    stop("[Import MSA validation: ] Characteristic object or path to metadata must be given.")
  }

  if(missing(prodata)) {
    stop("[Import MSA validation: ] ProData object or path to metadata must be given.")
  }

  if (method == "nested") {
    Object <- .NestedMSA(
      id = metadata[[1,2]],
      name = metadata[[2,2]],
      description = metadata[[3,2]],
      tolerance = as.numeric(metadata[[4,2]]),
      sigma = as.numeric(metadata[[5,2]]),
      alphaLim = as.numeric(metadata[[6,2]]),
      digits = as.numeric(metadata[[7,2]]),
      pro = pro,
      characteristic = characteristic,
      data = prodata
    )
  } else if (method == "crossed") {
    Object <- .CrossedMSA(
      id = metadata[[1,2]],
      name = metadata[[2,2]],
      description = metadata[[3,2]],
      tolerance = as.numeric(metadata[[4,2]]),
      sigma = as.numeric(metadata[[5,2]]),
      alphaLim = as.numeric(metadata[[6,2]]),
      digits = as.numeric(metadata[[7,2]]),
      pro = pro,
      characteristic = characteristic,
      data = prodata
    )
  } else if (method == "base") {
    Object <- .BaseMSA(
      id = metadata[[1,2]],
      name = metadata[[2,2]],
      description = metadata[[3,2]],
      tolerance = as.numeric(metadata[[4,2]]),
      sigma = as.numeric(metadata[[5,2]]),
      alphaLim = as.numeric(metadata[[6,2]]),
      digits = as.numeric(metadata[[7,2]]),
      pro = pro,
      characteristic = characteristic,
      data = prodata
    )
  } else {
    stop("[Import MSA validation: ] Method value must be, 'nested', 'crossed' or 'base'.")
  }


  return(Object)
}


