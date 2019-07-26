const async = require('async');
const mongoose = require('mongoose');
const base64Img = require("base64-img");
const fs = require('fs');
const Utils = require('../Utils/response')
const User = require('../models/User')
const Product = require('../models/Product')
const Account = require('../models/Account');
const Project = require('../models/Project');
const Idea = require('../models/Idea');
const Plan = require('../models/Plan');
const Requirement = require('../models/Requirement');
const WorkStreamModal = require('../models/WorkStream');
const Common = require('../common/common.js');
const Notification = require('../models/Notification');

var Excel = require('exceljs');

var product = module.exports = {

    /**
     * Add Product
     */
    addproduct: (req, res) => {
        if (req.decoded.id != undefined) {
            async.auto({
                findproduct: (callback) => {

                    Product.findOne({ adminId: req.decoded.adminId, productName: req.body.productName, isDeleted: false }, (err, response) => {
                        if (err) {
                            res.status(400).json(Utils.error(err))
                        } else if (response != null) {
                            res.status(400).json(Utils.error(Utils.PRODUCT_EXIST))
                        } else {
                            callback(null, response)
                        }
                    })
                },
                findLeadinUser: ['findproduct', (data, callback) => {
                    if (req.body.productLead != undefined) {
                        User.findOne({ _id: req.body.productLead }, (err, response) => {
                            if (err) {
                                res.status(400).json(Utils.error(err));
                            } else if (response == null) {
                                res.status(400).json(Utils.error(Utils.LEAD_NOT_EXIST));
                            } else {
                                callback(null, response)
                            }
                        })
                    } else {
                        callback(null, null)
                    }
                }],
                insertProduct: ['findLeadinUser', (data, callback) => {
                    if (data.findproduct == null) {
                        req.body.adminId = req.decoded.adminId
                        new Product(req.body).save((err, response) => {
                            if (err) {
                                res.status(400).json(Utils.error(err));
                            } else {
                                callback(null, response)
                            }
                        })
                    }
                }],
                sendNotification: ['insertProduct', (data, callback) => {
                    notificationObj = {
                        userId: req.decoded.id,
                        type: 3,
                        action: "added",
                        section: "product",
                        productId: data.insertProduct._id,
                        typeDetail: "Activity",
                        adminId: req.decoded.adminId,
                    }
                    var productInfo = data.insertProduct
                    Common.checkNotificationAccess(notificationObj, productInfo).then((response) => {
                        callback(null, response)
                    }).catch((error) => {
                        res.status(400).json(Utils.error(error));
                    })
                }],
                getAccountData: ['insertProduct', (data, callback) => {
                    Account.findOne({
                        userId: req.decoded.adminId
                    }).exec((err, accountresponse) => {
                        if (err) {
                            res.status(400).json(Utils.error(err));
                        } else {
                            callback(null, accountresponse)
                        }
                    })
                }],
                addPlan: ['getAccountData', (data, callback) => {

                    let stream = [{
                        title: 'Stream#1',
                        position: 'top'
                    }]
                    var timelineStartYear = new Date().getFullYear();
                    new Plan({ productId: data.insertProduct._id, workStreamArray: stream, timelineStartYear: timelineStartYear, configuration: data.getAccountData.configuration, startQuarter: data.getAccountData.startQuarter, endQuarter: data.getAccountData.endQuarter, startYear: data.getAccountData.startYear, endYear: data.getAccountData.endYear }).save((err, result) => {
                        if (err) {
                            res.status(400).json(Utils.error(err))
                        } else {
                            res.status(200).json(Utils.success(data.insertProduct, Utils.DATA_FETCH))
                        }

                    })
                }]
            })
        } else {
            res.status(401).json(Utils.authorisedError(Utils.AUTHENTICATION_FAILED))
        }
    },

    /**list product */

    productList: (req, res) => {
        if (req.decoded.id != undefined) {
            if (req.decoded.role == "Contributor" || req.decoded.role == "Stakeholder") {
                var criteria = {
                    productTeam: { $in: [mongoose.Types.ObjectId(req.decoded.id)] },
                    'isDeleted': false,
                    'default': false
                }
            } else {
                var criteria = {
                    'adminId': mongoose.Types.ObjectId(req.decoded.adminId),
                    'isDeleted': false,
                    'default': false
                }
            }

            Product.aggregate([{
                $match: criteria
            },
            {
                "$lookup": {
                    "from": "users",
                    "localField": "productLead",
                    "foreignField": "_id",
                    "as": "productLead"
                }
            },
            {
                "$unwind": {
                    path: "$productLead",
                    preserveNullAndEmptyArrays: true
                }
            },
            // {
            //     "$unwind": "$productLead"
            // },
            {
                "$lookup": {
                    "from": "accounts",
                    "localField": "adminId",
                    "foreignField": "userId",
                    "as": "adminId"
                }
            },
            {
                "$unwind": {
                    path: "$adminId",
                    preserveNullAndEmptyArrays: true
                }
            },
            // {
            //     "$unwind": "$adminId"
            // },
            {
                "$lookup": {
                    "from": "users",
                    "localField": "adminId.userId",
                    "foreignField": "_id",
                    "as": "adminId.userId"
                }
            },
            {
                "$unwind": {
                    path: "$adminId.userId",
                    preserveNullAndEmptyArrays: true
                }
            },
            // {
            //     "$unwind": "$adminId.userId"
            // },
            { $sort: { 'createdAt': -1 } }
            ]).exec((err, response) => {
                if (err) {
                    res.status(400).json(Utils.error(err))
                } else {
                    res.status(200).json(Utils.success(response, Utils.DATA_FETCH))
                }
            })
        } else {
            res.status(401).json(Utils.authorisedError(Utils.AUTHENTICATION_FAILED))
        }
    },

    /** Delete Product */
    productDelete: (req, res) => {
        Product.updateOne({ _id: req.params.id }, { $set: { isDeleted: true } }).exec((err, response) => {
            if (err) {
                res.status(400).json(Utils.error(err))
            } else {
                User.updateMany({ defaultProduct: req.params.id }, { $set: { defaultProduct: null } }).exec((err, response) => {
                    if (err) {
                        res.status(400).json(Utils.error(err))
                    } else {
                        res.status(200).json(Utils.success([], Utils.DELETE_PRODUCT))
                    }
                })

            }
        })
    },

    /** Add product tags */
    addProductTag: (req, res) => {
        Product.findOneAndUpdate({ _id: req.body.productId, isDeleted: false }, { $set: { tags: req.body.tags } }, { new: true }).exec((err, response) => {
            if (err) {
                res.status(400).json(Utils.error(err))
            } else if (response == null) {
                res.status(400).json(Utils.error(Utils.PRODUCT_NOT_FOUND));
            } else {
                res.status(200).json(Utils.success(response, Utils.TAG_ADDED))
            }
        })
    },

    /**
     * Edit product
     */
    productDetailUpdate: (req, res) => {

        if (req.decoded.id != undefined) {
            async.auto({
                uploadProductLogo: (callback) => {
                    let img = req.body.productLogo;
                    if (img != undefined && img.search("base64") > -1) {
                        var datetimestamp = Date.now();
                        var path = require("path").join(__dirname, "..", "public", "productLogo");
                        var filename = "product_" + datetimestamp;
                        base64Img.img(img, path, filename, function (err, filepath) {
                            if (err) {
                                res.status(400).json(Utils.error(err));
                            } else {
                                var lastslashindex = filepath.lastIndexOf("/");
                                var result = filepath.substring(lastslashindex + 1);
                                callback(null, result);
                            }
                        });
                    } else {
                        callback(null, null);
                    }
                },
                findLeadinUser: ['uploadProductLogo', (data, callback) => {
                    if (req.body.productLead != undefined) {
                        User.findOne({ _id: req.body.productLead }, (err, response) => {
                            if (err) {
                                res.status(400).json(Utils.error(err));
                            } else if (response == null) {
                                res.status(400).json(Utils.error(Utils.LEAD_NOT_EXIST));
                            } else {
                                callback(null, response)
                            }
                        })
                    } else {
                        callback(null, null)
                    }
                }],
                checkProduct: ['findLeadinUser', (data, callback) => {
                    if (req.body.productName != undefined) {
                        Product.findOne({ adminId: req.decoded.adminId, productName: req.body.productName }, (err, resposne) => {
                            if (err) {
                                res.status(400).json(Utils.error(err));
                            } else if (resposne != null) {
                                if (resposne.productName == req.body.productName && resposne._id != req.body.productId) {
                                    res.status(400).json(Utils.error(Utils.PRODUCT_EXIST));
                                } else {
                                    callback(null, 'true product')
                                }
                            } else {
                                callback(null, true)
                            }
                        })
                    } else {
                        callback(null, null)
                    }
                }],
                updateProduct: ['checkProduct', (data, callback) => {
                    if (req.body.vision != undefined) {
                        delete req.body.vision
                    }
                    if (data.uploadProductLogo != null) {
                        req.body.productLogo = data.uploadProductLogo;
                    }
                    Product.findOneAndUpdate({ _id: req.body.productId, isDeleted: false, adminId: req.decoded.adminId }, { $set: req.body }, { new: true, upsert: false }).exec((err, response) => {
                        if (err) {
                            res.status(400).json(Utils.error(err))
                        } else if (response === null) {
                            res.status(400).json(Utils.error(Utils.PRODUCT_NOT_FOUND));
                        } else {
                            callback(null, response)
                            //res.status(200).json(Utils.success(response, Utils.UPDATE_PRODUCT))
                        }
                    })
                }],
                sendNotification: ['updateProduct', (data, callback) => {
                    if (req.body.productLead != req.decoded.id) {

                        notificationObj = {
                            userId: req.decoded.id,
                            type: 3,
                            action: "added",
                            section: "product",
                            productId: data.updateProduct._id,
                            typeDetail: "Activity",
                            adminId: req.decoded.adminId,
                        }
                        var productInfo = data.updateProduct
                        Common.checkNotificationAccess(notificationObj, productInfo).then((response) => {

                            res.status(200).json(Utils.success(data.updateProduct, Utils.UPDATE_PRODUCT))

                        }).catch((error) => {
                            res.status(400).json(Utils.error(error));
                        })
                    } else {
                        res.status(200).json(Utils.success(data.updateProduct, Utils.UPDATE_PRODUCT))
                    }
                }],
            })
        } else {
            res.status(401).json(Utils.authorisedError(Utils.AUTHENTICATION_FAILED))
        }
    },

    /**
     * Update Product Vision
     */

    productUpateVision: (req, res) => {
        if (req.decoded.id != undefined) {
            async.auto({
                updateVision: (callback) => {
                    Product.findOneAndUpdate({ _id: req.body.productId, isDeleted: false, adminId: req.decoded.adminId }, { $set: req.body }, { new: true, upsert: false }).exec((err, response) => {
                        if (err) {
                            res.status(400).json(Utils.error(err))
                        } else if (response === null) {
                            res.status(400).json(Utils.error(Utils.PRODUCT_NOT_FOUND));
                        } else {
                            callback(null, response)
                            // res.status(200).json(Utils.success(response, Utils.UPDATE_PRODUCT))
                        }
                    })
                },
                sendNotification: ['updateVision', (data, callback) => {
                    notificationObj = {
                        userId: req.decoded.id,
                        type: 1,
                        action: "updated",
                        section: "product",
                        productId: req.body.productId,
                        typeDetail: "Strategy",
                        adminId: req.decoded.adminId
                    }
                    var productInfo = data.updateVision
                    Common.checkNotificationAccess(notificationObj, productInfo).then((response) => {

                        res.status(200).json(Utils.success(data.updateVision, Utils.DELETE_SUCCESS))

                    }).catch((error) => {
                        res.status(400).json(Utils.error(error));
                    })
                }]
            })
        } else {
            res.status(401).json(Utils.authorisedError(Utils.AUTHENTICATION_FAILED))
        }
    },

    /**
     * Add Customer segment and Competitor Profiles
     */
    addProductMarket: (req, res) => {
        if (req.decoded.id != undefined) {
            async.auto({
                uploadPhoto: (callback) => {
                    let img = req.body.photo;
                    if (img != undefined && img.search("base64") > -1) {
                        var datetimestamp = Date.now();
                        if (req.body.type == 'customer') {
                            var folderName = "customer"
                            var filename = "customer_" + datetimestamp;
                        } else {
                            var folderName = "competitor"
                            var filename = "competitor_" + datetimestamp;
                        }
                        var path = require("path").join(__dirname, "..", "public", folderName);
                        base64Img.img(img, path, filename, function (err, filepath) {
                            if (err) {
                                res.status(400).json(Utils.error(err));
                            } else {
                                var lastslashindex = filepath.lastIndexOf("/");
                                var result = filepath.substring(lastslashindex + 1);
                                callback(null, result);
                            }
                        });
                    } else {
                        callback(null, null);
                    }
                },
                addInformation: ['uploadPhoto', (data, callback) => {
                    if (data.uploadPhoto != null) {
                        req.body.photo = data.uploadPhoto
                    }
                    if (req.body.type == 'customer') {
                        var query = { 'customerSegments': req.body }
                    } else {
                        var query = { 'competitorProfiles': req.body }
                    }
                    Product.findOneAndUpdate({ _id: req.body.productId, isDeleted: false, adminId: req.decoded.adminId }, { $push: query }, { new: true, upsert: false }).exec((err, response) => {
                        if (err) {
                            res.status(400).json(Utils.error(err))
                        } else if (response === null) {
                            res.status(400).json(Utils.error(Utils.PRODUCT_NOT_FOUND));
                        } else {
                            callback(null, response)
                            //res.status(200).json(Utils.success(response, Utils.UPDATE_PRODUCT))
                        }
                    })
                }],
                sendNotification: ['addInformation', (data, callback) => {
                    notificationObj = {
                        userId: req.decoded.id,
                        type: 1,
                        action: "added",
                        section: "product",
                        productId: req.body.productId,
                        typeDetail: "Strategy",
                        adminId: req.decoded.adminId
                    }
                    var productInfo = data.addInformation
                    Common.checkNotificationAccess(notificationObj, productInfo).then((response) => {
                        res.status(200).json(Utils.success(data.addInformation, Utils.UPDATE_PRODUCT))
                    }).catch((error) => {
                        res.status(400).json(Utils.error(error));
                    })
                }]
            })
        } else {
            res.status(401).json(Utils.authorisedError(Utils.AUTHENTICATION_FAILED))
        }
    },

    /**
     * Edit Product Market (customer and competitor)
     */
    editProductMarket: (req, res) => {
        if (req.decoded.id != undefined) {
            async.auto({
                checkTitle: (callback) => {
                    if (req.body.type == 'customer') {
                        Product.findOne({ _id: req.body.productId, 'customerSegments.title': req.body.title }, { 'customerSegments.$': 1 }, (err, response) => {

                            if (err) {
                                res.status(400).json(Utils.error(err));
                            }
                            if (response == null) {
                                callback(null, true)
                            } else if (response.customerSegments[0].title == req.body.title && response.customerSegments[0]._id != req.body.id) {
                                res.status(400).json(Utils.error(Utils.TITLE_ALREADY_EXIST));
                            } else {
                                if (response.customerSegments[0].photo != undefined && req.body.photo != null) {
                                    callback(null, response.customerSegments[0].photo)
                                } else {
                                    callback(null, true)
                                }

                            }
                        });
                    } else {
                        Product.findOne({ _id: req.body.productId, 'competitorProfiles.title': req.body.title }, { 'competitorProfiles.$': 1 }, (err, response) => {
                            if (err) {
                                res.status(400).json(Utils.error(err));
                            }
                            if (response == null) {
                                callback(null, true)
                            } else if (response.competitorProfiles[0].title == req.body.title && response.competitorProfiles[0]._id != req.body.id) {
                                res.status(400).json(Utils.error(Utils.TITLE_ALREADY_EXIST));
                            } else {
                                if (response.competitorProfiles[0].photo != undefined && req.body.photo != null) {
                                    callback(null, response.competitorProfiles[0].photo)
                                } else {
                                    callback(null, true)
                                }
                            }
                        });
                    }
                },
                deletePreviousPhoto: ['checkTitle', (data, callback) => {
                    if (data.checkTitle != true) {
                        if (req.body.type == 'customer') {
                            var path = require("path").join(__dirname, "..", "public", 'customer');
                        } else {
                            var path = require("path").join(__dirname, "..", "public", 'competitor');
                        }
                        fs.unlink(path + '/' + data.checkTitle, (err) => {
                            if (err) {
                                callback(null, "no image found")
                            } else {
                                callback(null, 'fileDeleted')
                            }
                        });
                    } else {
                        callback(null, null)
                    }
                }],
                uploadPhoto: ['deletePreviousPhoto', (data, callback) => {
                    let img = req.body.photo;
                    if (img != undefined && img.search("base64") > -1) {
                        var datetimestamp = Date.now();
                        if (req.body.type == 'customer') {
                            var folderName = "customer"
                            var filename = "customer_" + datetimestamp;
                        } else {
                            var folderName = "competitor"
                            var filename = "competitor_" + datetimestamp;
                        }
                        var path = require("path").join(__dirname, "..", "public", folderName);

                        base64Img.img(img, path, filename, function (err, filepath) {
                            if (err) {
                                res.status(400).json(Utils.error(err));
                            } else {
                                var lastslashindex = filepath.lastIndexOf("/");
                                var result = filepath.substring(lastslashindex + 1);
                                callback(null, result);
                            }
                        });
                    } else {
                        callback(null, null);
                    }
                }],
                addInformation: ['uploadPhoto', (data, callback) => {
                    if (data.uploadPhoto != null) {
                        req.body.photo = data.uploadPhoto
                    }
                    if (req.body.type == 'customer') {
                        var query = { _id: req.body.productId, 'customerSegments._id': req.body.id, isDeleted: false, adminId: req.decoded.adminId }
                        var newBody = {
                            'customerSegments.$.title': req.body.title,
                            'customerSegments.$.importance': req.body.importance,
                            'customerSegments.$.tagline': req.body.tagline,
                            'customerSegments.$.discussion': req.body.discussion,
                            'customerSegments.$.needs': req.body.needs,
                            'customerSegments.$.photo': req.body.photo
                        }

                    } else {
                        var query = { _id: req.body.productId, 'competitorProfiles._id': req.body.id, isDeleted: false, adminId: req.decoded.adminId }
                        var newBody = {
                            'competitorProfiles.$.title': req.body.title,
                            'competitorProfiles.$.importance': req.body.importance,
                            'competitorProfiles.$.tagline': req.body.tagline,
                            'competitorProfiles.$.discussion': req.body.discussion,
                            'competitorProfiles.$.facts': req.body.facts,
                            'competitorProfiles.$.photo': req.body.photo
                        }

                    }
                    Product.findOneAndUpdate(query, { $set: newBody }, { new: true, upsert: false }).exec((err, response) => {
                        if (err) {
                            res.status(400).json(Utils.error(err))
                        } else if (response === null) {
                            res.status(400).json(Utils.error(Utils.PRODUCT_NOT_FOUND));
                        } else {
                            callback(null, response)
                            //res.status(200).json(Utils.success(response, Utils.UPDATE_PRODUCT))
                        }
                    })
                }],
                sendNotification: ['addInformation', (data, callback) => {
                    notificationObj = {
                        userId: req.decoded.id,
                        type: 1,
                        action: "updated",
                        section: "product",
                        productId: req.body.productId,
                        typeDetail: "Strategy",
                        adminId: req.decoded.adminId
                    }
                    var productInfo = data.addInformation
                    Common.checkNotificationAccess(notificationObj, productInfo).then((response) => {
                        res.status(200).json(Utils.success(data.addInformation, Utils.UPDATE_PRODUCT))

                    }).catch((error) => {
                        res.status(400).json(Utils.error(error));
                    })
                }]
            })
        } else {
            res.status(401).json(Utils.authorisedError(Utils.AUTHENTICATION_FAILED))
        }
    },


    /**
     * Get Product detail
     */
    getProductById: (req, res) => {
        if (req.decoded.id != undefined) {
            Common.checkExampleProduct(req.body.productId).then((response) => {
                if (response == true) {
                    var criteria = {
                        '_id': mongoose.Types.ObjectId(req.body.productId),
                        'isDeleted': false,
                        //'adminId': mongoose.Types.ObjectId(req.decoded.id)
                    }
                } else {
                    if (req.decoded.role == "Contributor" || req.decoded.role == "Stakeholder") {
                        var criteria = {
                            '_id': mongoose.Types.ObjectId(req.body.productId),
                            'isDeleted': false,
                            'productTeam': { '$in': [req.decoded.id] }
                        }
                    } else {
                        var criteria = {
                            '_id': mongoose.Types.ObjectId(req.body.productId),
                            'isDeleted': false,
                            'adminId': mongoose.Types.ObjectId(req.decoded.id)
                        }
                    }
                }
                Product.aggregate([{
                    $match: criteria
                },
                {
                    $lookup: {
                        from: 'accounts',
                        localField: 'adminId',
                        foreignField: 'userId',
                        as: 'adminId'
                    }
                },
                // { "$unwind": "$adminId" },
                {
                    "$unwind": {
                        path: "$adminId",
                        preserveNullAndEmptyArrays: true
                    }
                },
                {
                    $lookup: {
                        from: 'users',
                        localField: 'adminId.userId',
                        foreignField: '_id',
                        as: 'adminId.userId'
                    }
                },
                // { "$unwind": "$adminId.userId" },
                {
                    "$unwind": {
                        path: "$adminId.userId",
                        preserveNullAndEmptyArrays: true
                    }
                },
                {
                    $group: {
                        "_id": "$_id",
                        "adminId": { '$first': '$adminId' },
                        "productLead": { '$first': '$productLead' },
                        "productTeam": { '$first': '$productTeam' },
                        "productName": { '$first': '$productName' },
                        "productLogo": { '$first': '$productLogo' },
                        "strategyType": { '$first': '$strategyType' },
                        "vision": { '$first': '$vision' },
                        "competitorProfiles": { '$first': '$competitorProfiles' },
                        "customerSegments": { '$first': '$customerSegments' },
                        "objectives": { '$first': '$objectives' },
                        "focusArea": { '$first': '$focusArea' },
                        "description": { '$first': '$description' },
                        "tags": { '$first': '$tags' },
                        "focusArea": { '$first': '$focusArea' },
                        "isDeleted": { '$first': '$isDeleted' },
                        "deletedAT": { '$first': '$deletedAT' },
                        "createdAt": { '$first': '$createdAt' }
                    }
                }
                ]).exec((err, response) => {
                    if (err) {
                        res.status(400).json(Utils.error(err))
                    } else if (response.length == 0) {
                        res.status(400).json(Utils.error(Utils.PRODUCT_NOT_FOUND));
                    } else {
                        res.status(200).json(Utils.success(response[0], Utils.DATA_FETCH))
                    }
                });
            }).catch((err) => {
                res.status(400).json(Utils.error(err))
            })


        } else {
            res.status(401).json(Utils.authorisedError(Utils.AUTHENTICATION_FAILED))
        }
    },

    /**
     * Add Product Objectives
     */
    addProductObjectives: (req, res) => {
        if (req.decoded.id != undefined) {
            async.auto({
                addInformation: (callback) => {
                    Product.findOneAndUpdate({ _id: req.body.productId, isDeleted: false, adminId: req.decoded.adminId }, { $push: { objectives: req.body } }, { new: true, upsert: false }).exec((err, response) => {
                        if (err) {
                            res.status(400).json(Utils.error(err))
                        } else if (response === null) {
                            res.status(400).json(Utils.error(Utils.PRODUCT_NOT_FOUND));
                        } else {
                            callback(null, response)
                            // res.status(200).json(Utils.success(response, Utils.UPDATE_PRODUCT))
                        }
                    })
                },
                sendNotification: ['addInformation', (data, callback) => {
                    notificationObj = {
                        userId: req.decoded.id,
                        type: 1,
                        action: "added",
                        section: "product",
                        productId: req.body.productId,
                        typeDetail: "Strategy",
                        adminId: req.decoded.adminId
                    }
                    var productInfo = data.addInformation
                    Common.checkNotificationAccess(notificationObj, productInfo).then((response) => {

                        res.status(200).json(Utils.success(data.addInformation, Utils.UPDATE_PRODUCT))

                    }).catch((error) => {
                        res.status(400).json(Utils.error(error));
                    })
                }]
            })

        } else {
            res.status(401).json(Utils.authorisedError(Utils.AUTHENTICATION_FAILED))
        }
    },

    /**
     * Edit Product Objective
     */
    editProductObjective: (req, res) => {
        if (req.decoded.id != undefined) {
            async.auto({
                addInformation: (callback) => {
                    Product.findOneAndUpdate({ _id: req.body.productId, isDeleted: false, adminId: req.decoded.adminId, 'objectives._id': req.body.objectiveId }, {
                        $set: {
                            'objectives.$.title': req.body.title,
                            'objectives.$.shortName': req.body.shortName,
                            'objectives.$.kpi': req.body.kpi,
                            'objectives.$.importance': req.body.importance,
                            'objectives.$.discussion': req.body.discussion
                        }
                    }, { new: true, upsert: false }).exec((err, response) => {
                        if (err) {
                            res.status(400).json(Utils.error(err))
                        } else if (response === null) {
                            res.status(400).json(Utils.error(Utils.PRODUCT_NOT_FOUND));
                        } else {
                            callback(null, response)
                            // res.status(200).json(Utils.success(response, Utils.UPDATE_PRODUCT))
                        }
                    })
                },
                sendNotification: ['addInformation', (data, callback) => {
                    notificationObj = {
                        userId: req.decoded.id,
                        type: 1,
                        action: "updated",
                        section: "product",
                        productId: req.body.productId,
                        typeDetail: "Strategy",
                        adminId: req.decoded.adminId
                    }
                    var productInfo = data.addInformation
                    Common.checkNotificationAccess(notificationObj, productInfo).then((response) => {
                        res.status(200).json(Utils.success(data.addInformation, Utils.UPDATE_PRODUCT))

                    }).catch((error) => {
                        res.status(400).json(Utils.error(error));
                    })
                }]
            })
        } else {
            res.status(401).json(Utils.authorisedError(Utils.AUTHENTICATION_FAILED))
        }
    },

    /**
     * Add Product FocusArea
     */
    addProductFocusArea: (req, res) => {
        if (req.decoded.id != undefined) {
            async.auto({
                addInformation: (callback) => {
                    Product.findOneAndUpdate({ _id: req.body.productId, isDeleted: false, adminId: req.decoded.adminId }, { $push: { focusArea: req.body } }, { new: true, upsert: false }).exec((err, response) => {
                        if (err) {
                            res.status(400).json(Utils.error(err))
                        } else if (response === null) {
                            res.status(400).json(Utils.error(Utils.PRODUCT_NOT_FOUND));
                        } else {
                            callback(null, response)
                            // res.status(200).json(Utils.success(response, Utils.UPDATE_PRODUCT))
                        }
                    })
                },
                sendNotification: ['addInformation', (data, callback) => {
                    notificationObj = {
                        userId: req.decoded.id,
                        type: 1,
                        action: "added",
                        section: "product",
                        productId: req.body.productId,
                        typeDetail: "Strategy",
                        adminId: req.decoded.adminId
                    }
                    var productInfo = data.addInformation
                    Common.checkNotificationAccess(notificationObj, productInfo).then((response) => {

                        res.status(200).json(Utils.success(data.addInformation, Utils.UPDATE_PRODUCT))

                    }).catch((error) => {
                        res.status(400).json(Utils.error(error));
                    })
                }]
            })
        } else {
            res.status(401).json(Utils.authorisedError(Utils.AUTHENTICATION_FAILED))
        }
    },
    /**
     * Edit Product Focusarea
     */
    editProductfocusarea: (req, res) => {
        if (req.decoded.id != undefined) {
            async.auto({
                addInformation: (callback) => {
                    Product.findOneAndUpdate({
                        _id: req.body.productId,
                        isDeleted: false,
                        adminId: req.decoded.adminId,
                        'focusArea._id': req.body.focusareaId
                    }, {
                            $set: {
                                'focusArea.$.title': req.body.title,
                                'focusArea.$.shortName': req.body.shortName,
                                'focusArea.$.description': req.body.description,
                                'focusArea.$.assignTo': req.body.assignTo,
                                'focusArea.$.alignedObjective': req.body.alignedObjective
                            }
                        },

                        { new: true, upsert: false }).exec((err, response) => {
                            if (err) {
                                res.status(400).json(Utils.error(err))
                            } else if (response === null) {
                                res.status(400).json(Utils.error(Utils.PRODUCT_NOT_FOUND));
                            } else {
                                callback(null, response)
                                // res.status(200).json(Utils.success(response, Utils.UPDATE_PRODUCT))
                            }
                        })
                },
                sendNotification: ['addInformation', (data, callback) => {
                    notificationObj = {
                        userId: req.decoded.id,
                        type: 1,
                        action: "updated",
                        section: "product",
                        productId: req.body.productId,
                        typeDetail: "Strategy",
                        adminId: req.decoded.adminId
                    }
                    var productInfo = data.addInformation
                    Common.checkNotificationAccess(notificationObj, productInfo).then((response) => {

                        res.status(200).json(Utils.success(data.addInformation, Utils.UPDATE_PRODUCT))

                    }).catch((error) => {
                        res.status(400).json(Utils.error(error));
                    })
                }]
            })

        } else {
            res.status(401).json(Utils.authorisedError(Utils.AUTHENTICATION_FAILED))
        }
    },

    /**
     * Delete Product market and planning
     */
    deleteProductsubItem: (req, res) => {
        if (req.decoded.id != undefined) {
            async.auto({
                checkType: (callback) => {
                    if (req.body.type == 'customer') {
                        var query = { customerSegments: { _id: req.body.id } }
                        callback(null, query)
                    }
                    if (req.body.type == 'competitor') {
                        var query = { competitorProfiles: { _id: req.body.id } }
                        callback(null, query)
                    }

                    if (req.body.type == 'objective') {
                        var query = { objectives: { _id: req.body.id } }
                        callback(null, query)
                    }
                    if (req.body.type == 'focusarea') {
                        var query = { focusArea: { _id: req.body.id } }
                        callback(null, query)
                    }
                },
                deleteRecord: ['checkType', (data, callback) => {
                    var record = data.checkType
                    Product.findByIdAndUpdate({ _id: req.body.productId, isDeleted: false, adminId: req.decoded.adminId }, { $pull: record }, { new: true }).exec((err, response) => {
                        if (err) {
                            res.status(400).json(Utils.error(err))
                        } else {
                            callback(null, response)
                            // res.status(200).json(Utils.success(response, Utils.DELETE_SUCCESS))
                        }
                    })
                }],
                sendNotification: ['deleteRecord', (data, callback) => {
                    notificationObj = {
                        userId: req.decoded.id,
                        type: 1,
                        action: "updated",
                        section: "product",
                        productId: req.body.productId,
                        typeDetail: "Strategy",
                        adminId: req.decoded.adminId
                    }
                    var productInfo = data.deleteRecord
                    Common.checkNotificationAccess(notificationObj, productInfo).then((response) => {

                        res.status(200).json(Utils.success(data.deleteRecord, Utils.DELETE_SUCCESS))

                    }).catch((error) => {
                        res.status(400).json(Utils.error(error));
                    })
                }]
            })
        } else {
            res.status(401).json(Utils.authorisedError(Utils.AUTHENTICATION_FAILED))
        }
    },

    /**
     * List Of Product Name using admin Id
     */
    ListOfProductName: (req, res) => {
        Product.find({ adminId: req.body.adminId, isDeleted: false }, { productName: 1 }).exec((err, response) => {
            if (err) {
                res.status(400).json(Utils.error(err))
            } else {

                Account.findOne({ userId: req.body.adminId }).exec((err, accountdata) => {
                    if (err) {
                        res.status(400).json(Utils.error(err))
                    } else {
                        var result = {
                            statusCode: 200,
                            error: false,
                            success: true,
                            message: "data fetch successfully",
                            data: response,
                            adminId: accountdata
                        };
                        res.status(200).json(result)
                    }
                })
                // res.status(200).json(Utils.success(response, Utils.DATA_FETCH))
            }
        })
    },

    /**
     * Export (product, projects, ideas and strategy) product by id 
     */

    exportProductDetail: (req, res) => {
        let projectData1, objectives, ideaData1;
        if (req.decoded.id) {
            async.auto({
                productDetail: (callback) => {
                    Product.aggregate([{
                        $match: {
                            _id: mongoose.Types.ObjectId(req.body.productId),
                            isDeleted: false
                        }
                    },
                    {
                        $lookup: {
                            from: 'accounts',
                            localField: 'adminId',
                            foreignField: 'userId',
                            as: 'adminId'
                        }
                    },
                    { $unwind: "$adminId" },
                    ]).exec((err, response) => {
                        if (err) {
                            callback(err)
                        } else if (response.length == 0) {
                            res.status(400).json(Utils.error(Utils.PRODUCT_NOT_FOUND))
                        } else {
                            objectives = response[0].objectives;
                            callback(null, response[0])
                        }
                    })
                },
                findProjectData: ['productDetail', (data, callback) => {
                    Project.aggregate([{
                        $match: {
                            productId: mongoose.Types.ObjectId(req.body.productId),
                            isDeleted: false
                        }
                    },
                    { $addFields: { 'state_': { $cond: { if: { $eq: ['$state', 'Done'] }, then: 'done', else: '$state' } } } },
                    {
                        $lookup: {
                            from: 'products',
                            localField: 'productId',
                            foreignField: '_id',
                            as: 'productId'
                        }
                    },
                    { $unwind: "$productId" },

                    {
                        "$project": {
                            "totalObjective": { "$size": "$productId.objectives" },
                            "state_": 1,
                            "impactScore": 1,
                            "summary": 1,
                            "productId": 1,
                            "effortScore": 1,
                            "title": 1,
                            "summary": 1,
                            "tags": 1,
                            "businessCase": 1,
                            "state": 1,
                            "mustDoProject": 1,
                            "progress": 1,
                            "progressValue": 1,
                            "isDeleted": 1,
                            "text_module": 1,
                            "comments": 1,
                            "createdBy": 1,
                            "deletedAT": 1,
                            "createdAt": 1,
                            "focusArea": 1,
                            "attachments": 1,
                            "index": 1
                        }
                    },
                    {
                        "$addFields": {
                            "matchFocusArea": {
                                "$filter": {
                                    "input": "$productId.focusArea",
                                    "as": "el",
                                    "cond": {
                                        "$eq": [
                                            "$$el._id", "$focusArea"
                                        ]
                                    }
                                }
                            }
                        }
                    },
                    {
                        "$unwind": {
                            path: "$matchFocusArea",
                            preserveNullAndEmptyArrays: true
                        }
                    },
                    {
                        $lookup: {
                            from: 'users',
                            localField: 'createdBy',
                            foreignField: '_id',
                            as: 'createdBy'
                        }
                    },
                    { $unwind: "$createdBy" },
                    {
                        $group: {
                            "_id": "$_id",
                            "state_": { '$first': '$state_' },
                            "focusArea": { $push: "$matchFocusArea" },
                            "impactScore": { $first: "$impactScore" },
                            "totalObjective": { $first: "$totalObjective" },
                            "effortScore": { $first: "$effortScore" },
                            "title": { '$first': '$title' },
                            "summary": { '$first': '$summary' },
                            "tags": { '$first': '$tags' },
                            "effortScore": { '$first': '$effortScore' },
                            "businessCase": { '$first': '$businessCase' },
                            "state": { '$first': '$state' },
                            "mustDoProject": { '$first': '$mustDoProject' },
                            "progress": { '$first': '$progress' },
                            "progressValue": { '$first': '$progressValue' },
                            "isDeleted": { '$first': '$isDeleted' },
                            "text_module": { '$first': '$text_module' },
                            "comments": { '$first': '$comments' },
                            "firstName": { '$first': '$createdBy.firstName' },
                            "lastName": { '$first': '$createdBy.lastName' },
                            "deletedAT": { '$first': '$deletedAT' },
                            "createdAt": { '$first': '$createdAt' },
                            "attachments": { '$first': "$attachments" },
                            "index": { "$first": "$index" },
                            "productId": { "$first": "$productId" }
                        }
                    },
                    {
                        "$unwind": {
                            path: "$focusArea",
                            preserveNullAndEmptyArrays: true
                        }
                    },
                    {
                        "$unwind": {
                            path: "$productId.objectives",
                            preserveNullAndEmptyArrays: true
                        }
                    },

                    {
                        "$addFields": {
                            "matchRecords": {
                                "$filter": {
                                    "input": "$impactScore",
                                    "as": "el",
                                    "cond": {
                                        "$eq": [
                                            "$$el.objectiveId", "$productId.objectives._id"
                                        ]
                                    }
                                }
                            }
                        }
                    },
                    { $unwind: "$matchRecords" },

                    {
                        $group: {
                            "_id": "$_id",
                            "state_": { '$first': '$state_' },
                            "focusArea": { $first: "$focusArea._id" },
                            "focusAreaTitle": { $first: "$focusArea.title" },
                            "impactScore": { $push: "$matchRecords" },
                            "totalObjective": { $first: "$totalObjective" },
                            "effortScore": { $first: "$effortScore" },
                            "netValue": { '$first': "$netValue" },
                            "title": { '$first': '$title' },
                            "summary": { '$first': '$summary' },
                            "tags": { '$first': '$tags' },
                            "effortScore": { '$first': '$effortScore' },
                            "businessCase": { '$first': '$businessCase' },
                            "state": { '$first': '$state' },
                            "mustDoProject": { '$first': '$mustDoProject' },
                            "progress": { '$first': '$progress' },
                            "progressValue": { '$first': '$progressValue' },
                            "isDeleted": { '$first': '$isDeleted' },
                            "text_module": { '$first': '$text_module' },
                            "comments": { '$first': '$comments' },
                            "firstName": { '$first': '$firstName' },
                            "lastName": { '$first': '$lastName' },
                            "deletedAT": { '$first': '$deletedAT' },
                            "createdAt": { '$first': '$createdAt' },
                            "attachments": { '$first': "$attachments" },
                            "index": { "$first": "$index" }
                        }
                    },
                    {
                        $addFields: {
                            impactSum: {
                                "$reduce": {
                                    "input": "$impactScore",
                                    "initialValue": 0,
                                    "in": { "$add": ["$$value", "$$this.value"] }
                                }
                            }
                        }
                    },
                    {
                        $addFields: {
                            impact_average: { $divide: ["$impactSum", "$totalObjective"] }
                        }
                    },

                    {
                        $addFields: {
                            netValue: { $subtract: ["$impact_average", "$effortScore"] }
                        }
                    },
                    { $sort: { createdAt: -1 } },
                    ]).exec((err, response) => {
                        if (err) {
                            callback(err)
                        } else if (response.length == 0) {
                            callback(null, null)
                        } else {
                            projectData1 = response;
                            product.exportProjects(response, req.body.productId, data.productDetail.adminId).then((projects) => {
                                callback(null, projects)
                            }).catch((err) => {
                                callback(err)
                            })
                        }
                    })
                }],
                findIdeaDetail: ['findProjectData', (data, callback) => {
                    let workbook = data.findProjectData ? data.findProjectData.createWorkSheet.workbook : null;
                    let criteria = {
                        productId: req.body.productId,
                        isDeleted: false,
                        asProject: false,
                        asBacklog: false,
                        state: "Open"
                    }

                    Idea.aggregate([{
                        $match: {
                            productId: mongoose.Types.ObjectId(req.body.productId),
                            isDeleted: false,
                            asProject: false,
                            asBacklog: false,
                            state: "Open"
                        }
                    },
                    {
                        $lookup: {
                            from: 'users',
                            localField: 'createdBy',
                            foreignField: '_id',
                            as: 'createdBy'
                        }
                    },
                    {
                        "$unwind": {
                            path: "$createdBy",
                            preserveNullAndEmptyArrays: true
                        }
                    },
                    {
                        $lookup: {
                            from: 'departments',
                            localField: 'department',
                            foreignField: '_id',
                            as: 'department'
                        }
                    },
                    {
                        "$unwind": {
                            path: "$department",
                            preserveNullAndEmptyArrays: true
                        }
                    },
                    {
                        $group: {
                            "_id": "$_id",
                            "isDeleted": { '$first': '$isDeleted' },
                            "deletedAT": { '$first': '$deletedAT' },
                            "tags": { '$first': '$tags' },
                            "createdAt": { '$first': '$createdAt' },
                            "isUrgent": { '$first': '$isUrgent' },
                            "isFavourite": { '$first': '$isFavourite' },
                            "isDismiss": { '$first': '$isDismiss' },
                            "markAsRead": { '$first': '$markAsRead' },
                            "department": { '$first': '$department.name' },
                            "state": { '$first': '$state' },
                            "asProject": { '$first': '$asProject' },
                            "asBacklog": { '$first': '$asBacklog' },
                            "title": { '$first': '$title' },
                            "discussion": { '$first': '$discussion' },
                            "productId": { '$first': '$productId' },
                            "attachments": { '$first': '$attachments' },
                            "uniqueNumber": { '$first': '$uniqueNumber' },
                            "comments": { '$first': '$comments' },
                            "updatedAt": { '$first': '$updatedAt' },
                            "updatedBy": { '$first': '$updatedBy' },
                            "firstName": { '$first': '$createdBy.firstName' },
                            "lastName": { '$first': '$createdBy.lastName' },
                            "external": { '$first': '$external' },
                            "productId": { '$first': '$productId' }
                        }
                    },
                    { $sort: { 'createdAt': -1 } }
                    ]).exec((err, response) => {
                        if (err) {
                            callback(err)
                        } else if (response.length > 0) {
                            ideaData1 = response;
                            product.exportIdeas(response, req.body.productId, data.productDetail.adminId, workbook).then((result) => {
                                let dataToSend = result.createWorkSheet ? result.createWorkSheet.workbook : workbook;
                                callback(null, dataToSend)
                            }).catch((error) => {
                                callback(error)
                            });
                        } else {
                            callback(null, workbook)
                        }

                    })
                }],
                findProductDetail: ['findIdeaDetail', (data, callback) => {
                    let workbook = data.findIdeaDetail ? data.findIdeaDetail : null;
                    Product.aggregate([{
                        $match: {
                            _id: mongoose.Types.ObjectId(req.body.productId),
                            isDeleted: false,
                        }
                    },
                    {
                        $lookup: {
                            from: 'accounts',
                            localField: 'adminId',
                            foreignField: 'userId',
                            as: 'adminId'
                        }
                    },
                    {
                        "$unwind": {
                            path: "$adminId",
                            preserveNullAndEmptyArrays: true
                        }
                    },
                    ]).exec((err, productResponse) => {
                        if (err) {
                            callback(err)
                        } else if (productResponse.length == 0) {
                            callback(null, null)
                        } else {
                            product.exportProductDetailData(productResponse[0], req.body.productId, data.productDetail.adminId, projectData1, objectives, ideaData1, workbook).then((response) => {
                                callback(null, response)
                            }).catch((error) => {
                                callback(error)
                            })
                        }
                    })
                }]
            }, (err, results) => {
                if (err) { } else {
                    var fileName = 'productReport_' + req.body.productId + '.xlsx'

                    res.status(200).json({ file: fileName, message: 'Data Exported successfully' })
                }
            })
        } else {
            res.status(401).json(Utils.authorisedError(Utils.AUTHENTICATION_FAILED))
        }
    },

    /**
     * export projects
     * @param {projectData} projectData 
     * @param {productId} productId 
     */
    exportProjects(projectData, productId, adminId) {
        return new Promise((resolve, reject) => {
            var fileName = 'productReport_' + productId + '.xlsx'
            async.auto({
                createWorkSheet: (callback) => {
                    var fileName = 'productReport_' + productId + '.xlsx'
                    let path = require("path").join(__dirname, '..', 'public', 'productLogo', fileName);

                    initWorkbook(path)
                        .then(workbook => {
                            sheet = workbook.addWorksheet('projects');
                            sheet.columns = [
                                { header: 'Title', key: 'Title', width: 20 },
                                { header: 'State', key: 'State', width: 20 },
                                { header: 'Theme', key: 'Theme', width: 20 },
                                { header: 'ProgressValue', key: 'ProgressValue', width: 20 },
                                { header: 'Summary', key: 'Summary', width: 20 },
                                { header: 'BusinessCase', key: 'BusinessCase', width: 20 },
                                { header: 'EffortScore', key: 'EffortScore', width: 20 },
                                // { header: 'Tags', key: 'Tags', width: 20 },
                                { header: 'Impact_average', key: 'Impact_average', width: 20 },
                                { header: 'CreatedBy', key: 'CreatedBy', width: 20 },
                                { header: 'NetValue', key: 'NetValue', width: 20 },

                            ];
                            async.forEachOf(projectData, function (project, key, cb) {
                                var tagDetail = JSON.stringify(project.tags)
                                var createdBy = project.firstName + ' ' + project.lastName;
                                var rounded = Math.round(project.impact_average * 10) / 10;
                                sheet.addRow({
                                    'Title': project.title,
                                    'State': project.state,
                                    'Theme': project.focusAreaTitle,
                                    'ProgressValue': project.progressValue,
                                    'Summary': project.summary,
                                    'BusinessCase': project.businessCase,
                                    'EffortScore': project.effortScore,
                                    // 'Tags': tagDetail,
                                    'Impact_average': rounded,
                                    'CreatedBy': createdBy,
                                    'NetValue': Math.round(project.netValue * 10) / 10,

                                });
                                cb();
                            })

                            workbook.xlsx.writeFile(path)
                                .then(function () {
                                    let dataToSend = {
                                        message: Utils.DATA_FETCH,
                                        file: fileName,
                                        workbook: workbook
                                    };
                                    callback(null, dataToSend);
                                });
                        }).catch(err => {
                            console.error(err);
                            cb(err);
                        });

                }
            }, (error, result) => {
                if (error) {
                    reject(error)
                } else {
                    resolve(result)
                }
            })
        })
    },

    /**
     * export ideas
     * @param {ideaData} ideaData 
     * @param {productId} productId 
     */
    exportIdeas(ideaData, productId, adminId, workbook) {

        return new Promise((resolve, reject) => {
            async.auto({
                createWorkSheet: (callback) => {

                    var fileName = 'productReport_' + productId + '.xlsx'
                    let path = require("path").join(__dirname, '..', 'public', 'productLogo', fileName);
                    if (workbook) {
                        // initWorkbook(path)
                        // .then(workbook => {
                        sheet = workbook.addWorksheet('Idea');
                        sheet.columns = [
                            { header: 'Title', key: 'Title', width: 20 },
                            // { header: 'Source', key: 'Source', width: 20 },
                            { header: 'CreatedAt', key: 'CreatedAt', width: 20 },
                            { header: 'Discussion', key: 'Discussion', width: 20 },
                            { header: 'IsUrgent', key: 'IsUrgent', width: 20 },
                            { header: 'IsFavourite', key: 'IsFavourite', width: 20 },
                            // { header: 'Tags', key: 'Tags', width: 20 },
                            // { header: 'index', key: 'index', width: 20 },
                            { header: 'CreatedBy', key: 'CreatedBy', width: 20 },
                            { header: 'State', key: 'State', width: 20 },

                        ];

                        async.forEachOf(ideaData, function (idea, key, cb) {

                            var tagDetail = JSON.stringify(idea.tags)
                            var createdBy = idea.firstName + ' ' + idea.lastName;
                            sheet.addRow({
                                'Title': idea.title,
                                // 'Source': idea.department,
                                'CreatedAt': idea.createdAt,
                                'Discussion': idea.discussion,
                                'IsUrgent': idea.isUrgent,
                                'IsFavourite': idea.isFavourite,
                                // 'Tags': idea.tags,
                                // 'index': idea.uniqueNumber,
                                'CreatedBy': createdBy,
                                'State': idea.state,

                            });
                            cb();
                        })

                        workbook.xlsx.writeFile(path)
                            .then(function () {
                                let dataToSend = {
                                    message: Utils.DATA_FETCH,
                                    file: fileName,
                                    workbook: workbook
                                };
                                callback(null, dataToSend);
                            })
                            .catch(err => {
                                console.error(err);
                                cb(err);
                            });
                        // }).catch(err => {
                        //     console.error(err);
                        //     cb(err);
                        // });
                    } else {
                        initWorkbook(path)
                            .then(workbook => {
                                sheet = workbook.addWorksheet('Idea');
                                sheet.columns = [
                                    { header: 'Title', key: 'Title', width: 20 },
                                    // { header: 'Source', key: 'Source', width: 20 },
                                    { header: 'CreatedAt', key: 'CreatedAt', width: 20 },
                                    { header: 'Discussion', key: 'Discussion', width: 20 },
                                    { header: 'IsUrgent', key: 'IsUrgent', width: 20 },
                                    { header: 'IsFavourite', key: 'IsFavourite', width: 20 },
                                    // { header: 'Tags', key: 'Tags', width: 20 },
                                    // { header: 'index', key: 'index', width: 20 },
                                    { header: 'CreatedBy', key: 'CreatedBy', width: 20 },
                                    { header: 'State', key: 'State', width: 20 },

                                ];

                                async.forEachOf(ideaData, function (idea, key, cb) {
                                    var tagDetail = JSON.stringify(idea.tags)
                                    var createdBy = idea.firstName + ' ' + idea.lastName;
                                    sheet.addRow({
                                        'Title': idea.title,
                                        // 'Source': idea.department,
                                        'CreatedAt': idea.createdAt,
                                        'Discussion': idea.discussion,
                                        'IsUrgent': idea.isUrgent,
                                        'IsFavourite': idea.isFavourite,
                                        // 'Tags': idea.tags,
                                        // 'index': idea.uniqueNumber,
                                        'CreatedBy': idea.firstName + ' ' + idea.lastName,
                                        'State': idea.state,

                                    });
                                    cb();
                                })

                                workbook.xlsx.writeFile(path)
                                    .then(function () {
                                        let dataToSend = {
                                            message: Utils.DATA_FETCH,
                                            file: fileName,
                                            workbook: workbook
                                        };
                                        callback(null, dataToSend);
                                    });
                            }).catch(err => {
                                console.error(err);
                                cb(err);
                            });
                    }
                }
            }, (error, result) => {
                if (error) {
                    reject(error)
                } else {
                    resolve(result)
                }
            })
        })
    },

    exportProductDetailData(productData, productId, adminId, projectData1, objectives, ideaData1, workbook) {
        return new Promise((resolve, reject) => {
            async.auto({
                customerSheet: (callback) => {
                    var customerData = productData.customerSegments;
                    if (customerData.length > 0) {

                        var fileName = 'productReport_' + productId + '.xlsx'
                        let path = require("path").join(__dirname, '..', 'public', 'productLogo', fileName);
                        if (workbook) {
                            // initWorkbook(path)
                            // .then(workbook => {
                            sheet = workbook.addWorksheet('customers');
                            sheet.columns = [
                                { header: 'Title', key: 'Title', width: 20 },
                                { header: 'Importance', key: 'Importance', width: 20 },
                                { header: 'Tagline', key: 'Tagline', width: 20 },
                                { header: 'Discussion', key: 'Discussion', width: 20 },
                                { header: 'Needs', key: 'Needs', width: 20 }

                            ];

                            async.forEachOf(customerData, function (customer, key, cb) {
                                sheet.addRow({
                                    'Title': customer.title,
                                    'Importance': customer.importance,
                                    'Tagline': customer.tagline,
                                    'Discussion': customer.discussion,
                                    'Needs': customer.needs
                                });
                                cb();
                            })

                            workbook.xlsx.writeFile(path)
                                .then(function () {
                                    let dataToSend = {
                                        message: Utils.DATA_FETCH,
                                        file: fileName
                                    };
                                    callback(null, dataToSend);
                                })
                                .catch(err => {
                                    console.error(err);
                                    callback(err);
                                })
                        } else {
                            initWorkbook(path)
                                .then(workbook => {
                                    sheet = workbook.addWorksheet('customers');
                                    sheet.columns = [
                                        { header: 'Title', key: 'Title', width: 20 },
                                        { header: 'Importance', key: 'Importance', width: 20 },
                                        { header: 'Tagline', key: 'Tagline', width: 20 },
                                        { header: 'Discussion', key: 'Discussion', width: 20 },
                                        { header: 'Needs', key: 'Needs', width: 20 }

                                    ];

                                    async.forEachOf(customerData, function (customer, key, cb) {
                                        sheet.addRow({
                                            'Title': customer.title,
                                            'Importance': customer.importance,
                                            'Tagline': customer.tagline,
                                            'Discussion': customer.discussion,
                                            'Needs': customer.needs
                                        });
                                        cb();
                                    })

                                    workbook.xlsx.writeFile(path)
                                        .then(function () {
                                            let dataToSend = {
                                                message: Utils.DATA_FETCH,
                                                file: fileName
                                            };
                                            callback(null, dataToSend);
                                        })
                                        .catch(err => {
                                            callback(err);
                                        })
                                }).catch(err => {
                                    cb(err);
                                });
                        }


                    } else {
                        callback(null, null)
                    }
                },
                competitorSheet: ['customerSheet', (data, callback) => {
                    var competitorData = productData.competitorProfiles;
                    if (competitorData.length > 0) {

                        var fileName = 'productReport_' + productId + '.xlsx'
                        let path = require("path").join(__dirname, '..', 'public', 'productLogo', fileName);

                        if (workbook) {
                            sheet = workbook.addWorksheet('competitor');
                            sheet.columns = [
                                { header: 'Title', key: 'Title', width: 20 },
                                { header: 'Importance', key: 'Importance', width: 20 },
                                { header: 'Tagline', key: 'Tagline', width: 20 },
                                { header: 'Discussion', key: 'Discussion', width: 20 },
                                { header: 'Facts', key: 'Facts', width: 20 }

                            ];

                            async.forEachOf(competitorData, function (competitor, key, cb) {
                                sheet.addRow({
                                    'Title': competitor.title,
                                    'Importance': competitor.importance,
                                    'Tagline': competitor.tagline,
                                    'Discussion': competitor.discussion,
                                    'Facts': competitor.needs
                                });
                                cb();
                            })

                            workbook.xlsx.writeFile(path)
                                .then(function () {
                                    let dataToSend = {
                                        message: Utils.DATA_FETCH,
                                        file: fileName
                                    };
                                    callback(null, dataToSend);
                                })
                                .catch(err => {
                                    console.error(err);
                                    callback(err);
                                })
                        } else {
                            initWorkbook(path)
                                .then(workbook => {
                                    sheet = workbook.addWorksheet('Competitor');
                                    sheet.columns = [
                                        { header: 'Title', key: 'Title', width: 20 },
                                        { header: 'Importance', key: 'Importance', width: 20 },
                                        { header: 'Tagline', key: 'Tagline', width: 20 },
                                        { header: 'Discussion', key: 'Discussion', width: 20 },
                                        { header: 'Needs', key: 'Needs', width: 20 }

                                    ];

                                    async.forEachOf(competitorData, function (competitor, key, cb) {
                                        sheet.addRow({
                                            'Title': competitor.title,
                                            'Importance': competitor.importance,
                                            'Tagline': competitor.tagline,
                                            'Discussion': competitor.discussion,
                                            'Needs': competitor.needs
                                        });
                                        cb();
                                    })

                                    workbook.xlsx.writeFile(path)
                                        .then(function () {
                                            let dataToSend = {
                                                message: Utils.DATA_FETCH,
                                                file: fileName
                                            };
                                            callback(null, dataToSend);
                                        })
                                        .catch(err => {
                                            console.error(err);
                                            callback(err);
                                        })
                                }).catch(err => {
                                    console.error(err);
                                    cb(err);
                                });
                        }

                    } else {
                        callback(null, null)
                    }
                }],
                objectiveSheet: ['competitorSheet', (data, callback) => {
                    var objectivesData = productData.objectives;
                    if (objectivesData.length > 0) {

                        var fileName = 'productReport_' + productId + '.xlsx'
                        let path = require("path").join(__dirname, '..', 'public', 'productLogo', fileName);

                        if (workbook) {
                            sheet = workbook.addWorksheet(adminId.objectives_label);
                            sheet.columns = [
                                { header: 'Title', key: 'Title', width: 20 },
                                { header: 'Importance', key: 'Importance', width: 20 },
                                { header: 'ShortName', key: 'ShortName', width: 20 },
                                { header: 'Discussion', key: 'Discussion', width: 20 },
                                { header: 'Kpi', key: 'Kpi', width: 20 }

                            ];

                            async.forEachOf(objectivesData, function (objectives, key, cb) {
                                sheet.addRow({
                                    'Title': objectives.title,
                                    'Importance': objectives.importance,
                                    'ShortName': objectives.shortName,
                                    'Discussion': objectives.discussion,
                                    'Kpi': objectives.kpi
                                });
                                cb();
                            })

                            workbook.xlsx.writeFile(path)
                                .then(function () {
                                    let dataToSend = {
                                        message: Utils.DATA_FETCH,
                                        file: fileName
                                    };
                                    callback(null, dataToSend);
                                })
                                .catch(err => {
                                    console.error(err);
                                    callback(err);
                                })
                        } else {

                            initWorkbook(path)
                                .then(workbook => {
                                    sheet = workbook.addWorksheet(adminId.objectives_label);
                                    sheet.columns = [
                                        { header: 'Title', key: 'Title', width: 20 },
                                        { header: 'Importance', key: 'Importance', width: 20 },
                                        { header: 'ShortName', key: 'ShortName', width: 20 },
                                        { header: 'Discussion', key: 'Discussion', width: 20 },
                                        { header: 'Kpi', key: 'Kpi', width: 20 }

                                    ];

                                    async.forEachOf(objectivesData, function (objectives, key, cb) {
                                        sheet.addRow({
                                            'Title': objectives.title,
                                            'Importance': objectives.importance,
                                            'ShortName': objectives.shortName,
                                            'Discussion': objectives.discussion,
                                            'Kpi': objectives.kpi
                                        });
                                        cb();
                                    })

                                    workbook.xlsx.writeFile(path)
                                        .then(function () {
                                            let dataToSend = {
                                                message: Utils.DATA_FETCH,
                                                file: fileName
                                            };
                                            callback(null, dataToSend);
                                        })
                                        .catch(err => {
                                            console.error(err);
                                            callback(err);
                                        })
                                }).catch(err => {
                                    console.error(err);
                                    cb(err);
                                });
                        }


                    } else {
                        callback(null, null)
                    }
                }],
                focusAreaSheet: ['objectiveSheet', (data, callback) => {
                    var focusAreaData = productData.focusArea;
                    if (focusAreaData.length > 0) {

                        var fileName = 'productReport_' + productId + '.xlsx'
                        let path = require("path").join(__dirname, '..', 'public', 'productLogo', fileName);

                        if (workbook) {
                            sheet = workbook.addWorksheet(adminId.theme_label);
                            sheet.columns = [
                                { header: 'Title', key: 'Title', width: 20 },
                                // { header: 'Importance', key: 'Importance', width: 20 },
                                { header: 'ShortName', key: 'ShortName', width: 20 },
                                { header: 'Description', key: 'Description', width: 20 }

                            ];

                            async.forEachOf(focusAreaData, function (focusArea, key, cb) {
                                sheet.addRow({
                                    'Title': focusArea.title,
                                    // 'Importance': focusArea.importance,
                                    'ShortName': focusArea.shortName,
                                    'Description': focusArea.description
                                });
                                cb();
                            })

                            workbook.xlsx.writeFile(path)
                                .then(function () {
                                    let dataToSend = {
                                        message: Utils.DATA_FETCH,
                                        file: fileName
                                    };
                                    callback(null, dataToSend);
                                })
                                .catch(err => {
                                    console.error(err);
                                    callback(err);
                                })
                        } else {

                            initWorkbook(path)
                                .then(workbook => {
                                    sheet = workbook.addWorksheet(adminId.theme_label);
                                    sheet.columns = [
                                        { header: 'Title', key: 'Title', width: 20 },
                                        // { header: 'Importance', key: 'Importance', width: 20 },
                                        { header: 'ShortName', key: 'ShortName', width: 20 },
                                        { header: 'Description', key: 'Description', width: 20 }

                                    ];

                                    async.forEachOf(focusAreaData, function (focusArea, key, cb) {
                                        sheet.addRow({
                                            'Title': focusArea.title,
                                            // 'Importance': focusArea.importance,
                                            'ShortName': focusArea.shortName,
                                            'Description': focusArea.description
                                        });
                                        cb();
                                    })

                                    workbook.xlsx.writeFile(path)
                                        .then(function () {
                                            let dataToSend = {
                                                message: Utils.DATA_FETCH,
                                                file: fileName
                                            };
                                            callback(null, dataToSend);
                                        })
                                        .catch(err => {
                                            console.error(err);
                                            callback(err);
                                        })
                                }).catch(err => {
                                    console.error(err);
                                    cb(err);
                                });
                        }

                    } else {
                        callback(null, null)
                    }
                }],
                visionSheet: ['focusAreaSheet', (data, callback) => {
                    // if (productData.vision) {
                    var fileName = 'productReport_' + productId + '.xlsx'
                    let path = require("path").join(__dirname, '..', 'public', 'productLogo', fileName);
                    if (productData.vision) {
                        if (workbook) {
                            sheet = workbook.addWorksheet('Product Vision');
                            sheet.columns = [
                                { header: 'Product Vision', key: 'Product Vision', width: 50 }
                                // { header: 'Tags', key: 'Tags', width: 30 }
                            ];

                            sheet.addRow({
                                'Product Vision': productData.vision
                            });
                            // async.forEach(productData.tags, (tag, cb) => {
                            //     sheet.addRow({
                            //         'Tags': tag
                            //     });
                            //     cb();
                            // }, (e, r) => {
                            workbook.xlsx.writeFile(path)
                                .then(function () {
                                    let dataToSend = {
                                        message: Utils.DATA_FETCH,
                                        file: fileName
                                    };
                                    callback(null, dataToSend);
                                })
                                .catch(err => {
                                    console.error(err);
                                    callback(err);
                                })
                            // })


                        } else {

                            initWorkbook(path)
                                .then(workbook => {
                                    sheet = workbook.addWorksheet('Product Vision');
                                    sheet.columns = [
                                        { header: 'Product Vision', key: 'Product Vision', width: 30 }
                                        // { header: 'Tags', key: 'Tags', width: 30 }
                                    ];

                                    sheet.addRow({
                                        'Product Vision': productData.vision
                                    });
                                    // async.forEach(productData.tags, (tag, cb) => {
                                    //     sheet.addRow({
                                    //         'Tags': tag
                                    //     });
                                    //     cb();
                                    // }, (e, r) => {
                                    workbook.xlsx.writeFile(path)
                                        .then(function () {
                                            let dataToSend = {
                                                message: Utils.DATA_FETCH,
                                                file: fileName
                                            };
                                            callback(null, dataToSend);
                                        })
                                        .catch(err => {
                                            console.error(err);
                                            callback(err);
                                        })
                                    // })
                                })
                        }

                    } else {
                        callback(null, null)
                    }
                }],

                productTagsSheet: ['visionSheet', (data, callback) => {
                    // if (productData.vision) {
                    var fileName = 'productReport_' + productId + '.xlsx'
                    let path = require("path").join(__dirname, '..', 'public', 'productLogo', fileName);
                    if (productData.tags && productData.tags.length > 0) {
                        if (workbook) {
                            sheet = workbook.addWorksheet('Product Tags');
                            sheet.columns = [
                                // { header: 'Vision', key: 'Vision', width: 50 }
                                { header: 'Product Tags', key: 'Product Tags', width: 30 }
                            ];

                            // sheet.addRow({
                            //     'Vision': productData.vision
                            // });
                            async.forEach(productData.tags, (tag, cb) => {
                                sheet.addRow({
                                    'Product Tags': tag
                                });
                                cb();
                            }, (e, r) => {
                                workbook.xlsx.writeFile(path)
                                    .then(function () {
                                        let dataToSend = {
                                            message: Utils.DATA_FETCH,
                                            file: fileName
                                        };
                                        callback(null, dataToSend);
                                    })
                                    .catch(err => {
                                        console.error(err);
                                        callback(err);
                                    })
                            })


                        } else {

                            initWorkbook(path)
                                .then(workbook => {
                                    sheet = workbook.addWorksheet('Product Tags');
                                    sheet.columns = [
                                        // { header: 'Vision', key: 'Vision', width: 30 }
                                        { header: 'Product Tags', key: 'Product Tags', width: 30 }
                                    ];

                                    // sheet.addRow({
                                    //     'Vision': productData.vision
                                    // });
                                    async.forEach(productData.tags, (tag, cb) => {
                                        sheet.addRow({
                                            'Product Tags': tag
                                        });
                                        cb();
                                    }, (e, r) => {
                                        workbook.xlsx.writeFile(path)
                                            .then(function () {
                                                let dataToSend = {
                                                    message: Utils.DATA_FETCH,
                                                    file: fileName
                                                };
                                                callback(null, dataToSend);
                                            })
                                            .catch(err => {
                                                console.error(err);
                                                callback(err);
                                            })
                                    })
                                })
                        }
                    } else {
                        callback(null, null)
                    }
                }],

                projectObjectivesSheet: ['productTagsSheet', (data, callback) => {
                    if (projectData1 && projectData1.length > 0) {

                        var fileName = 'productReport_' + productId + '.xlsx'
                        let path = require("path").join(__dirname, '..', 'public', 'productLogo', fileName);

                        if (workbook) {
                            sheet = workbook.addWorksheet('Project Objectives');
                            sheet.columns = [
                                { header: 'Project Title', key: 'Project Title', width: 20 },
                                { header: 'Objective Name', key: 'Objective Name', width: 20 },
                                { header: 'Objective Value', key: 'Objective Value', width: 20 }

                            ];

                            async.forEach(projectData1, function (project, cb) {
                                async.forEach(project.impactScore, (obj, cb1) => {
                                    var results = objectives.filter(function (entry) {
                                        return entry._id.toString() == obj.objectiveId.toString();
                                    });
                                    sheet.addRow({
                                        'Project Title': project.title,
                                        'Objective Name': results[0].title,
                                        'Objective Value': obj.value
                                    });

                                    cb1();
                                }, (e, r) => {
                                    cb();
                                })
                            }, (err, res) => {
                                workbook.xlsx.writeFile(path)
                                    .then(function () {
                                        let dataToSend = {
                                            message: Utils.DATA_FETCH,
                                            file: fileName
                                        };
                                        callback(null, dataToSend);
                                    })
                                    .catch(err => {
                                        console.error(err);
                                        callback(err);
                                    })
                            })

                        } else {
                            if (projectData1 && projectData1.length > 0) {
                                initWorkbook(path)
                                    .then(workbook => {
                                        sheet = workbook.addWorksheet('Project Objectives');
                                        sheet.columns = [
                                            { header: 'Project Title', key: 'Project Title', width: 20 },
                                            { header: 'Objective Name', key: 'Objective Name', width: 20 },
                                            { header: 'Objective Value', key: 'Objective Value', width: 20 }

                                        ];

                                        async.forEach(projectData1, function (project, cb) {
                                            async.forEach(project.impactScore, (obj, cb1) => {
                                                var results = objectives.filter(function (entry) {
                                                    return entry._id.toString() == obj.objectiveId.toString();
                                                });
                                                sheet.addRow({
                                                    'Project Title': project.title,
                                                    'Objective Name': results[0].title,
                                                    'Objective Value': obj.value
                                                });

                                                cb1();
                                            }, (e, r) => {
                                                cb();
                                            })
                                        }, (err, res) => {
                                            workbook.xlsx.writeFile(path)
                                                .then(function () {
                                                    let dataToSend = {
                                                        message: Utils.DATA_FETCH,
                                                        file: fileName
                                                    };
                                                    callback(null, dataToSend);
                                                })
                                                .catch(err => {
                                                    console.error(err);
                                                    callback(err);
                                                })
                                        })
                                    }).catch(err => {
                                        console.error(err);
                                        cb(err);
                                    });
                            } else {
                                callback(null, null)
                            }
                        }

                    } else {
                        callback(null, null)
                    }
                }],

                projectTagsSheet: ['projectObjectivesSheet', (data, callback) => {
                    if (projectData1 && projectData1.length > 0) {
                        // console.
                        var fileName = 'productReport_' + productId + '.xlsx'
                        let path = require("path").join(__dirname, '..', 'public', 'productLogo', fileName);

                        if (workbook) {
                            sheet = workbook.addWorksheet('Project Tags');
                            sheet.columns = [
                                { header: 'Project Title', key: 'Project Title', width: 20 },
                                { header: 'Tag', key: 'Tag', width: 20 },
                            ];
                            let count = 0;

                            async.forEach(projectData1, function (project, cb) {
                                if (project.tags && project.tags.length == 0) {
                                    count = count + 1;
                                }
                                async.forEach(project.tags, (tag, cb1) => {
                                    sheet.addRow({
                                        'Project Title': project.title,
                                        'Tag': tag,
                                    });
                                    cb1();
                                }, (e, r) => {
                                    cb();
                                })
                            }, (err, res) => {
                                if (projectData1.length != count) {
                                    workbook.xlsx.writeFile(path)
                                        .then(function () {
                                            let dataToSend = {
                                                message: Utils.DATA_FETCH,
                                                file: fileName
                                            };
                                            callback(null, dataToSend);
                                        })
                                        .catch(err => {
                                            console.error(err);
                                            callback(err);
                                        })
                                } else {
                                    callback(null, null);
                                }
                            })

                        } else {
                            if (projectData1 && projectData1.length > 0) {
                                initWorkbook(path)
                                    .then(workbook => {
                                        sheet = workbook.addWorksheet('Project Objectives');
                                        sheet.columns = [
                                            { header: 'Project Title', key: 'Project Title', width: 20 },
                                            { header: 'Tag', key: 'Tag', width: 20 },

                                        ];
                                        let count = 0;
                                        async.forEach(projectData1, function (project, cb) {
                                            if (project.tags && project.tags.length == 0) {
                                                count = count + 1;
                                            }
                                            async.forEach(project.tags, (tag, cb1) => {
                                                sheet.addRow({
                                                    'Project Title': project.title,
                                                    'Tag': tag,
                                                });
                                                cb1();
                                            }, (e, r) => {
                                                cb();
                                            })
                                        }, (err, res) => {
                                            if (projectData1.length != count) {
                                                workbook.xlsx.writeFile(path)
                                                    .then(function () {
                                                        let dataToSend = {
                                                            message: Utils.DATA_FETCH,
                                                            file: fileName
                                                        };
                                                        callback(null, dataToSend);
                                                    })
                                                    .catch(err => {
                                                        console.error(err);
                                                        callback(err);
                                                    })
                                            } else {
                                                callback(null, null);
                                            }
                                        })
                                    }).catch(err => {
                                        console.error(err);
                                        cb(err);
                                    });
                            } else {
                                callback(null, null)
                            }
                        }

                    } else {
                        callback(null, null)
                    }
                }],

                ideasTagsSheet: ['projectTagsSheet', (data, callback) => {
                    if (ideaData1 && ideaData1.length > 0) {

                        var fileName = 'productReport_' + productId + '.xlsx'
                        let path = require("path").join(__dirname, '..', 'public', 'productLogo', fileName);

                        if (workbook) {
                            sheet = workbook.addWorksheet('Idea Tags');
                            sheet.columns = [
                                { header: 'Idea Title', key: 'Idea Title', width: 20 },
                                { header: 'Tag', key: 'Tag', width: 20 },
                            ];
                            let count = 0;
                            async.forEach(ideaData1, function (idea, cb) {
                                if (idea.tags && idea.tags.length == 0) {
                                    count = count + 1;
                                }
                                async.forEach(idea.tags, (tag, cb1) => {
                                    sheet.addRow({
                                        'Idea Title': idea.title,
                                        'Tag': tag,
                                    });
                                    cb1();
                                }, (e, r) => {
                                    cb();
                                })
                            }, (err, res) => {
                                if (ideaData1.length != count) {
                                    workbook.xlsx.writeFile(path)
                                        .then(function () {
                                            let dataToSend = {
                                                message: Utils.DATA_FETCH,
                                                file: fileName
                                            };
                                            callback(null, dataToSend);
                                        })
                                        .catch(err => {
                                            console.error(err);
                                            callback(err);
                                        })
                                } else {
                                    callback(null, null)
                                }
                            })

                        } else {
                            if (ideaData1 && ideaData1.length > 0) {
                                initWorkbook(path)
                                    .then(workbook => {
                                        sheet = workbook.addWorksheet('Idea Tags');
                                        sheet.columns = [
                                            { header: 'Idea Title', key: 'Idea Title', width: 20 },
                                            { header: 'Tag', key: 'Tag', width: 20 },

                                        ];
                                        let count = 0;

                                        async.forEach(ideaData1, function (idea, cb) {
                                            if (idea.tags && idea.tags.length == 0) {
                                                count = count + 1;
                                            }
                                            async.forEach(idea.tags, (tag, cb1) => {
                                                sheet.addRow({
                                                    'Idea Title': idea.title,
                                                    'Tag': tag,
                                                });
                                                cb1();
                                            }, (e, r) => {
                                                cb();
                                            })
                                        }, (err, res) => {
                                            if (ideaData1.length != count) {
                                                workbook.xlsx.writeFile(path)
                                                    .then(function () {
                                                        let dataToSend = {
                                                            message: Utils.DATA_FETCH,
                                                            file: fileName
                                                        };
                                                        callback(null, dataToSend);
                                                    })
                                                    .catch(err => {
                                                        console.error(err);
                                                        callback(err);
                                                    })
                                            } else {
                                                callback(null, null)
                                            }
                                        })
                                    }).catch(err => {
                                        console.error(err);
                                        cb(err);
                                    });
                            } else {
                                callback(null, null)
                            }
                        }

                    } else {
                        callback(null, null)
                    }
                }],

            }, (err, results) => {
                if (err) {
                    reject(err)
                } else {
                    resolve(results)
                }
            })
        })
    },

    /**
     * Example product data
     */
    exampleProductDetail: (req, res) => {
        if (req.decoded.id) {
            Product.findOne({ default: true }).select('_id, productName').exec((err, response) => {
                if (err) {
                    res.status(400).json(Utils.error(err))
                } else {
                    res.status(200).json(Utils.success(response, Utils.DATA_FETCH))
                }
            })
        } else {
            res.status(401).json(Utils.authorisedError(Utils.AUTHENTICATION_FAILED))
        }
    }

}

function initWorkbook(path) {
    var workbook = new Excel.Workbook();
    var worksheet;
    return new Promise(function (resolve, reject) {
        try {
            if (fs.existsSync(path)) {
                // load the Excel workbook but do nothing
                workbook.xlsx.readFile(path).then(function () {
                    workbook.eachSheet(function (worksheet, sheetId) {

                        workbook.removeWorksheet(sheetId)
                    });
                    resolve(workbook);
                });
            } else {

                resolve(workbook);
            }
        } catch (err) {
            reject();
        }
    });

}