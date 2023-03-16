/* eslint-disable react/display-name */
import * as React from "react";
import { Provider, TextArea, Form, FormButton, Dropdown, Menu, Image, Loader, Checkbox, Alert } from "@fluentui/react-northstar";
import { useState, useEffect } from "react";
import { useTeams } from "msteams-react-base-component";
import { app, dialog } from "@microsoft/teams-js";
/**
 * Implementation of the yoteam Message Extension Task Module page  , PASSWORD, USERID
 */

interface badgeInterface {
        "header": String,
        "sys_id": String,
        "image": URL,
        "user_recieve": String
};

const userId = process.env.REACT_APP_INSTANCE_USERID;
const Password = process.env.REACT_APP_INSTANCE_PASSWORD;
const encoder = new TextEncoder();
const credentials = `${userId}:${Password}`;
const encodedCredentials = encoder.encode(credentials);
const base64Credentials = btoa(String.fromCharCode.apply(null, encodedCredentials));

async function fetchData(id: string): Promise<any> {
    const response = await fetch(`https://${process.env.REACT_APP_INSTANCE}/api/now/table/${id}`, {
        method: "GET",
        headers: {
            Accept: "application/json",
            "Content-Type": "application/json",
            Authorization: `Basic ${base64Credentials}`
        }
    }).then(response => response.json())
        .catch(error => {
            console.error("Error fetching data:", error);
        });
    return response.result;
}

async function userCriteria(id: string): Promise<any> {
    const criteriaCall = await fetch(`https://${process.env.REACT_APP_INSTANCE}/api/now/table/${id}`, {
        method: "GET",
        headers: {
            Accept: "application/json",
            "Content-Type": "application/json",
            Authorization: `Basic ${base64Credentials}`
        }
    })
        .then(response => response.json())
        .then(async criteria => {
            if (criteria.result.length > 0) {
                return await fetch(`https://${process.env.REACT_APP_INSTANCE}/api/now/table/user_criteria?sysparm_query=sys_id=${criteria.result[0].user_criteria.value}&sysparm_fields=user`, {
                    method: "GET",
                    headers: {
                        Accept: "application/json",
                        "Content-Type": "application/json",
                        Authorization: `Basic ${base64Credentials}`
                    }
                }).then(response => response.json())
                    .catch(error => {
                        console.error("Error fetching data:", error);
                    });
            }
        });
    return criteriaCall;
}

export const YoteamMessageExtensionAction = () => {

    const [{ context, inTeams, theme }] = useTeams();
    const [activeIndex, setActiveIndex] = useState(0);

    const [inputUsersList, setInputUsersList] = useState<any>();
    const [inputRecValues, setInputRecValues] = useState<any>();
    const [inputBadges, setInputBadges] = useState<any>();
    const [inputSkills, setinputSkills] = useState<any>();

    const [isLoading, setIsLoading] = useState(true);
    const [alertVisible, setAlertVisible] = useState(true);
    const [useSkill, setUseSkill] = useState("");

    const [description, setDescription] = useState("");
    const [user, setUser] = useState<any>();
    const [badge, setBadge] = useState<any>();
    const [recValue, setRecValue] = useState<any>();
    const [skill, setSkill] = useState<any>();
    const [sysId, setSysId] = useState("");

    const [maxRecognitions, setmaxRecognitions] = useState(0);
    const [maxSkills, setmaxSkills] = useState(0);

    const [note, setNote] = useState("");
    const [checked, setChecked] = useState(false);
    const [category, setCategory] = useState<any>();
    const [image, setImage] = useState<any>();

    const [inputcategoryList, setInputCategoryList] = useState<any>();
    const [inputimagesList, setInputImagesList] = useState<any>();

    const [alertStatus, setAlertStatus] = useState("");
    const [alertContent, setAlertContent] = useState("");

    const handleWishSubmit = () => {
        if (!user || !image || note === "") {
            setAlertVisible(true);
            setAlertStatus("Danger");
            setAlertContent("Please fill all the mandatory fields");
            setTimeout(() => {
                setAlertVisible(false);
            }, 2000);
            return;
        }
        console.log(checked);
        const jsonBody = JSON.stringify({
            ecard_to: "1",
            personal_notes: note,
            u_recognized_by: sysId,
            recipient: user.sys_id,
            u_send_later: "false",
            u_state: "u_sent",
            u_ecard: image.sys_id,
            private: checked
        });
        console.log(jsonBody);
        const requestOptions = {
            method: "POST",
            headers: {
                "Content-Type": "application/json",
                Authorization: `Basic ${base64Credentials}`
            },
            body: jsonBody
        };
        setIsLoading(true);
        fetch( `https://${process.env.REACT_APP_INSTANCE}/api/now/table/x_figin_kudos_ecards?sysparm_fields=number`, requestOptions)
            .then((response) => response.json())
            .then((res) => {
                console.log(res.result.number);
                setAlertVisible(true);
                setAlertStatus("Success");
                setAlertContent(`Your Wish Request ${res.result.number} has been submitted successfully`);
                setTimeout(() => {
                    setAlertVisible(false);
                    dialog.url.submit();
                }, 2000);
                setIsLoading(false);
            })
            .catch(error => {
                console.log(error);
                setAlertVisible(true);
                setAlertStatus("Danger");
                setAlertContent("Error While Sending Wish Request");
            });
    };

    const handlePraiseSubmit = () => {
        if (!user || !badge || !recValue || recValue.length === 0 || description === "") {
            setAlertVisible(true);
            setAlertStatus("Danger");
            setAlertContent("Please fill all the mandatory fields");
            setTimeout(() => {
                setAlertVisible(false);
            }, 2000);
            return;
        }
        if (useSkill === "true" && (skill === undefined || skill.length === 0)) {
            setAlertVisible(true);
            setAlertStatus("Danger");
            setAlertContent("Please fill all the mandatory fields");
            setTimeout(() => {
                setAlertVisible(false);
            }, 2000);
            return;
        }
        if ((badge as badgeInterface).user_recieve !== undefined) {
            if (!badge.user_recieve.includes(user.sys_id)) {
                setAlertVisible(true);
                setAlertStatus("Danger");
                setAlertContent(`${badge.header} badge cannot be given to the particular user`);
                setTimeout(() => {
                    setAlertVisible(false);
                }, 2000);
                return setUser("");
            }
        }
        const recogsysIds = recValue.map(item => item.sys_id);
        let skillsysIds;
        let jsonBody;
        if (useSkill === "true") {
            skillsysIds = skill.map(item => item.sys_id);
            jsonBody = JSON.stringify({
                badge: badge.sys_id,
                description,
                kudos_from: sysId,
                employee: user.sys_id,
                type: "1",
                u_skill: skillsysIds.toString(),
                recognized_values: recogsysIds.toString()
            });
        } else {
            jsonBody = JSON.stringify({
                badge: badge.sys_id,
                description,
                kudos_from: sysId,
                employee: user.sys_id,
                type: "1",
                recognized_values: recogsysIds.toString()
            });
        }
        const requestOptions = {
            method: "POST",
            headers: {
                "Content-Type": "application/json",
                Authorization: `Basic ${base64Credentials}`
            },
            body: jsonBody
        };
        setIsLoading(true);
        fetch( `https://${process.env.REACT_APP_INSTANCE}/api/now/table/x_figin_kudos_recognition?sysparm_fields=number`, requestOptions)
            .then((response) => response.json())
            .then((res) => {
                console.log(res.result.number);
                setAlertVisible(true);
                setAlertStatus("Success");
                setAlertContent(`Your Praise Request ${res.result.number} has been submitted successfully`);
                setTimeout(() => {
                    setAlertVisible(false);
                    dialog.url.submit();
                }, 2000);
                setIsLoading(false);
            })
            .catch(error => {
                console.log(error);
                setAlertVisible(true);
                setAlertStatus("Danger");
                setAlertContent("Error While Sending Recognition Request");
            });
    };

    useEffect(() => {
        if (context?.user?.userPrincipalName !== undefined) {
            console.log(context);
            console.log(process.env.REACT_APP_INSTANCE);
            const fetchallData = async() => {
                try {
                    const usersysId = await fetchData(`sys_user?sysparm_query=email=${context?.user?.userPrincipalName}&sysparm_fields=sys_id`);
                    setSysId(usersysId[0].sys_id);
                    let userListcall = await fetchData(`sys_user?sysparm_query=active=true^roles=x_figin_kudos.user^ORroles=x_figin_kudos.admin^ORroles=admin^sys_id!=${usersysId[0].sys_id}&sysparm_fields=sys_id,name`);
                    userListcall = userListcall.map(obj => {
                        return { header: obj.name, sys_id: obj.sys_id };
                    });
                    setInputUsersList(userListcall);
                    let recognizedValuesCall = await fetchData("x_figin_kudos_recognized_value?sysparm_query=active=true&sysparm_fields=sys_id,name");
                    recognizedValuesCall = recognizedValuesCall.map(obj => {
                        return { header: obj.name, sys_id: obj.sys_id };
                    });
                    setInputRecValues(recognizedValuesCall);
                    const maxRecognitionsCall = await fetchData("sys_properties?name=x_figin_kudos.max_values_selected");
                    setmaxRecognitions(maxRecognitionsCall[0].value);
                    const displayskillCall = await fetchData("sys_properties?name=x_figin_kudos.skillsForRecognition");
                    setUseSkill(displayskillCall[0].value);
                    let badgesCall = await fetchData("x_figin_kudos_badge?sysparm_query=type=1&sysparm_fields=sys_id,title,image");
                    badgesCall = badgesCall.map(async obj => {
                        const userNominate = await userCriteria(`x_figin_kudos_m2m_user_criteri_badges_nominate?sysparm_query=badge=${obj.sys_id}&sysparm_fields=user_criteria`);
                        if (userNominate !== undefined) {
                            if (userNominate.result[0].user.includes(usersysId[0].sys_id)) {
                                const imgUrl = await fetch(`https://${process.env.REACT_APP_INSTANCE}/api/now/attachment/${obj.image}/file`, {
                                    method: "GET",
                                    headers: {
                                        Accept: "application/json",
                                        "Content-Type": "application/json",
                                        Authorization: `Basic ${base64Credentials}`
                                    }
                                }).then(async (response) => {
                                    const blob = await response.blob();
                                    const objectUrl = URL.createObjectURL(blob);
                                    return objectUrl;
                                }).catch(error => {
                                    console.log(error);
                                    setAlertVisible(true);
                                    setAlertStatus("Danger");
                                    setAlertContent("Error While Fetch Attachment Image Data");
                                }); ;
                                return { header: obj.title, sys_id: obj.sys_id, image: imgUrl };
                            } else { /* empty */ }
                        } else {
                            const imgUrl = await fetch(`https://${process.env.REACT_APP_INSTANCE}/api/now/attachment/${obj.image}/file`, {
                                method: "GET",
                                headers: {
                                    Accept: "application/json",
                                    "Content-Type": "application/json",
                                    Authorization: `Basic ${base64Credentials}`
                                }
                            }).then(async (response) => {
                                const blob = await response.blob();
                                const objectUrl = URL.createObjectURL(blob);
                                return objectUrl;
                            }).catch(error => {
                                console.log(error);
                                setAlertVisible(true);
                                setAlertStatus("Danger");
                                setAlertContent("Error While Fetch Attachment Image Data");
                            }); ;
                            return { header: obj.title, sys_id: obj.sys_id, image: imgUrl };
                        }
                    });
                    let badgesList = await Promise.all(badgesCall);
                    badgesList = badgesList.filter(Boolean);
                    setInputBadges(badgesList);
                    let categoryCall = await fetchData("x_figin_kudos_ecard_category?sysparm_query=ORDERBYname&sysparm_fields=sys_id,name");
                    categoryCall = categoryCall.map(obj => {
                        return { header: obj.name, sys_id: obj.sys_id };
                    });
                    setInputCategoryList(categoryCall);
                    setCategory(categoryCall[0]);
                    setIsLoading(false);
                    badgesList = badgesList.map(async obj => {
                        const userRecieve = await userCriteria(`x_figin_kudos_m2m_user_criteri_badges?sysparm_query=badge=${obj.sys_id}&sysparm_fields=user_criteria`);
                        if (userRecieve !== undefined) {
                            return { ...obj, user_recieve: userRecieve.result[0].user };
                        } else {
                            return obj;
                        }
                    });
                    setInputBadges(await Promise.all(badgesList));
                } catch (error) {
                    console.log(error);
                    setAlertVisible(true);
                    setAlertStatus("Danger");
                    setAlertContent("Error While Fetch Form Data");
                }
            };
            fetchallData();
        }
    }, [context?.user?.userPrincipalName]);

    useEffect(() => {
        const fetchaskillData = async() => {
            try {
                const maxSkillsCall = await fetchData("sys_properties?name=x_figin_kudos.max_skills_selected");
                setmaxSkills(maxSkillsCall[0].value);
                let SkillListCall = await fetchData("cmn_skill_m2m_category?sysparm_display_value=true&sysparm_query=category=ecaf8f98dbcb8150df4f230848961928^skill.active=true^ORDERBYskill&sysparm_fields=sys_id,skill");
                SkillListCall = SkillListCall.map(obj => {
                    return { header: obj.skill.display_value, sys_id: obj.sys_id };
                });
                setinputSkills(SkillListCall);
            } catch (error) {
                console.log(error);
                setAlertVisible(true);
                setAlertStatus("Danger");
                setAlertContent("Error While Fetch Skills Data");
            }
        };
        if (useSkill === "true") {
            fetchaskillData();
        }
    }, [useSkill]);

    useEffect(() => {
        const fetchimageData = async() => {
            try {
                setIsLoading(true);
                let imagesCall = await fetchData(`x_figin_kudos_ecard_images?sysparm_query=active=true^ecard_category.sys_id=${category.sys_id}&sysparm_fields=sys_id,u_title,image`);
                imagesCall = imagesCall.map(async obj => {
                    const imgUrl = await fetch(`https://${process.env.REACT_APP_INSTANCE}/api/now/attachment/${obj.image}/file`, {
                        method: "GET",
                        headers: {
                            Accept: "application/json",
                            "Content-Type": "application/json",
                            Authorization: `Basic ${base64Credentials}`
                        }
                    }).then(async (response) => {
                        const blob = await response.blob();
                        const objectUrl = URL.createObjectURL(blob);
                        return objectUrl;
                    }).catch(error => {
                        console.error("Error fetching data:", error);
                    }); ;
                    return { header: obj.u_title, sys_id: obj.sys_id, image: imgUrl };
                });
                setInputImagesList(await Promise.all(imagesCall));
                setImage("");
                setIsLoading(false);
            } catch (error) {
                console.log(error);
                setAlertVisible(true);
                setAlertStatus("Danger");
                setAlertContent("Error While Fetch Image Data");
            }
        };
        if (category !== undefined) {
            fetchimageData();
        }
    }, [category]);

    const Wish = (
        <Form
            onSubmit={handleWishSubmit}
            style={{
                alignItems: "center",
                width: "100%",
                padding: "10px",
                justifyContent: "flex-start",
                overflow: "scroll"
            }}
        >
            <div className="drop-width">
                <span id="wishlabel">Wish<span className="astrick">*</span> :</span>
                <Dropdown
                    items={inputUsersList}
                    placeholder="Select users..."
                    search
                    fluid
                    aria-labelledby="wishlabel"
                    noResultsMessage="We couldn't find any matches."
                    style={{ marginTop: "10px" }}
                    value={user}
                    onChange={(e, data) => {
                        if (data.value) {
                            setUser(data.value);
                        }
                    }}
                /></div>
            <div className="drop-width">
                <span id="notelabel">Personal Note<span className="astrick">*</span> :</span>
                <TextArea
                    placeholder="Type here..."
                    maxLength={500}
                    aria-labelledby="notelabel"
                    fluid
                    value={note}
                    style={{
                        height: "100px",
                        marginTop: "10px"
                    }}
                    onChange={(e, data) => {
                        if (data?.value) {
                            setNote(data.value);
                        }
                    }}
                    name="personal note"
                /></div>
            <div className="drop-width">
                <Checkbox checked={checked} label="Private" onClick={(e, data) => {
                    if (data) {
                        setChecked(data.checked);
                    }
                }}
                />
            </div>
            <div className="drop-width">
                <span id="wishlabel">Wish Category<span className="astrick">*</span> :</span>
                <Dropdown
                    items={inputcategoryList}
                    placeholder="Select category..."
                    fluid
                    aria-labelledby="wishlabel"
                    style={{ marginTop: "10px" }}
                    value={category}
                    onChange={(e, data) => {
                        if (data.value) {
                            setCategory(data.value);
                        }
                    }}
                /></div>
            <div className="drop-width">
                <span id="imagelabel">Wish Image<span className="astrick">*</span> :</span>
                <Dropdown
                    items={inputimagesList}
                    placeholder="Select image..."
                    fluid
                    className="imgDropdown"
                    aria-labelledby="imagelabel"
                    style={{ marginTop: "10px" }}
                    value={image}
                    onChange={(e, data) => {
                        if (data.value) {
                            setImage(data.value);
                        }
                    }}
                />
                {image && (<Image
                    fluid
                    src={image?.image}
                />)}
            </div>
            <FormButton primary content="Submit" />

        </Form>
    );

    const Praise = (
        <Form
            onSubmit={handlePraiseSubmit}
            style={{
                alignItems: "center",
                width: "100%",
                padding: "10px",
                justifyContent: "flex-start",
                overflow: "scroll"
            }}
        >
            <div className="drop-width">
                <span id="userlabel">Praise<span className="astrick">*</span> :</span>
                <Dropdown
                    items={inputUsersList}
                    placeholder="Select users..."
                    search
                    fluid
                    aria-labelledby="userlabel"
                    noResultsMessage="We couldn't find any matches."
                    style={{ marginTop: "10px" }}
                    value={user}
                    onChange={(e, data) => {
                        if (data.value) {
                            setUser(data.value);
                        }
                    }}
                /></div>
            <div className="drop-width">
                <span id="badgeLabel">Badge<span className="astrick">*</span> :</span>
                <Dropdown
                    fluid
                    search
                    aria-labelledby="badgeLabel"
                    items={inputBadges}
                    placeholder="Select badges..."
                    noResultsMessage="We couldn't find any matches."
                    style={{ marginTop: "10px" }}
                    value={badge}
                    onChange={(e, data) => {
                        if (data) {
                            setBadge(data.value);
                        }
                    }}
                /></div>
            <div className="drop-width">
                <span id="valuesLabel">Recognized Values<span className="astrick">*</span> :</span>
                <Dropdown
                    fluid
                    aria-labelledby="valuesLabel"
                    search
                    multiple
                    items={inputRecValues}
                    placeholder="Select values..."
                    noResultsMessage="We couldn't find any matches."
                    style={{ marginTop: "10px" }}
                    value={recValue}
                    onChange={(e, data) => {
                        if (data) {
                            if ((data.value as Array<Object>).length <= maxRecognitions) {
                                setRecValue(data.value);
                            }
                        }
                    }}
                /></div>
            {useSkill === "true" && (
                <div className="drop-width">
                    <span id="Skillslabel">Skills<span className="astrick">*</span> :</span>
                    <Dropdown
                        fluid
                        search
                        multiple
                        aria-labelledby="Skillslabel"
                        items={inputSkills}
                        placeholder="Select skills..."
                        noResultsMessage="We couldn't find any matches."
                        style={{ marginTop: "10px" }}
                        value={skill}
                        onChange={(e, data) => {
                            if (data) {
                                if ((data.value as Array<Object>).length <= maxSkills) {
                                    setSkill(data.value);
                                }
                            }
                        }}
                    /></div>
            )}
            <div className="drop-width">
                <span id="Descriptionlabel">Description<span className="astrick">*</span> :</span>
                <TextArea
                    placeholder="Type here..."
                    maxLength={500}
                    aria-labelledby="Descriptionlabel"
                    fluid
                    value={description}
                    style={{
                        height: "100px",
                        marginTop: "10px"
                    }}
                    onChange={(e, data) => {
                        if (data?.value) {
                            setDescription(data.value);
                        }
                    }}
                    name="description"
                /></div>
            <FormButton primary content="Submit" />
        </Form>
    );

    useEffect(() => {
        if (inTeams === true) {
            app.notifySuccess();
        }
    }, [inTeams]);

    const panes = [
        {
            menuItem: { key: "praise", content: "Praise" },
            render: () => Praise
        },
        {
            menuItem: { key: "wish", content: "Wish" },
            render: () => Wish
        }
    ];

    const handleMenuItemClick = (e, { index }) => {
        setUser("");
        setActiveIndex(index);
    };

    if (isLoading) {
        return (
            <><Provider theme={theme}>
                {alertVisible && alertStatus === "Success" && (
                    <Alert success content={alertContent} className="alertPad" />
                )}
                {alertVisible && alertStatus === "Danger" && (
                    <Alert danger content={alertContent} className="alertPad" />
                )}
            </Provider>
            <Provider theme={theme} style={{ height: "100%", display: "flex", justifyContent: "center" }}>
                <Loader label="Loading..." />
            </Provider></>
        );
    }

    return (
        <><Provider theme={theme}>
            {alertVisible && alertStatus === "Success" && (
                <Alert success content={alertContent} className="alertPad" />
            )}
            {alertVisible && alertStatus === "Danger" && (
                <Alert danger content={alertContent} className="alertPad" />
            )}
        </Provider>
        <Provider theme={theme} style={{ position: "relative", height: "90%" }}>
            <Menu
                underlined
                primary
                items={panes.map((pane, index) => ({
                    ...pane.menuItem,
                    active: activeIndex === index,
                    index,
                    onClick: handleMenuItemClick
                }))}
                style={{
                    border: "none",
                    margin: "10px auto"
                }}
            />
            {panes[activeIndex].render()}
        </Provider>
        </>
    );
};
